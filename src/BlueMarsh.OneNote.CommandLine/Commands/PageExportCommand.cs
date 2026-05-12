using System.CommandLine;
using System.Text;
using System.Xml.Linq;
using BlueMarsh.OneNote.CommandLine.Export;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

internal static class PageExportCommand
{
    public static Command Create()
    {
        var refArg = new Argument<string>("ref")
        {
            Description = "Page, section, section group, or notebook (name, ID, or path)",
        };
        var outputDirOption = new Option<DirectoryInfo?>("--output-dir", "--out", "-o")
        {
            Description = "Output directory for exported files (defaults to current directory)",
        };
        outputDirOption.AcceptLegalFilePathsOnly();
        var rawOption = new Option<bool>("--raw")
        {
            Description = "Export raw OneNote XML instead of Markdown",
        };

        var command = new Command("export", "Export pages to Markdown or XML files")
        {
            refArg,
            outputDirOption,
            rawOption,
        };

        command.SetAction(parseResult =>
        {
            var refString = parseResult.GetValue(refArg)!;
            var outputDir = parseResult.GetValue(outputDirOption)
                ?? new DirectoryInfo(Directory.GetCurrentDirectory());
            var raw = parseResult.GetValue(rawOption);

            if (!outputDir.Exists)
                outputDir.Create();

            using var oneNote = new OneNoteApplication();

            var resolved = OneNoteRef.Resolve(oneNote, refString);
            if (resolved is null)
            {
                Console.Error.WriteLine($"'{refString}' not found.");
                return;
            }

            OneNotePageToMarkdownConverter? converter = null;
            if (!raw)
            {
                converter = new OneNotePageToMarkdownConverter(
                    warn: msg => Console.Error.WriteLine($"Warning: {msg}"));
            }

            switch (resolved.NodeType)
            {
                case HierarchyNodeType.Page:
                    ExportSinglePage(oneNote, converter, resolved, outputDir);
                    break;

                case HierarchyNodeType.Section:
                    ExportSection(oneNote, converter, resolved.Id, resolved.Name, outputDir);
                    break;

                case HierarchyNodeType.SectionGroup:
                case HierarchyNodeType.Notebook:
                    ExportContainer(oneNote, converter, resolved, outputDir);
                    break;
            }
        });

        return command;
    }

    private static void ExportPage(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter? converter,
        string pageId,
        string pageName,
        DirectoryInfo outputDir)
    {
        var pageXml = oneNote.GetPageContent(pageId);

        string content;
        string extension;
        if (converter is not null)
        {
            content = converter.Convert(pageXml);
            extension = ".md";
        }
        else
        {
            content = XDocument.Parse(pageXml).ToString();
            extension = ".xml";
        }

        var filePath = Path.Combine(outputDir.FullName, SanitizeFileName(pageName) + extension);
        File.WriteAllText(filePath, content, Encoding.UTF8);
        Console.WriteLine(filePath);
    }

    private static void ExportSinglePage(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter? converter,
        ResolvedRef pageRef,
        DirectoryInfo outputDir)
    {
        ExportPage(oneNote, converter, pageRef.Id, pageRef.Name, outputDir);
    }

    private static void ExportSection(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter? converter,
        string sectionId,
        string sectionName,
        DirectoryInfo outputDir)
    {
        var sectionDir = outputDir.CreateSubdirectory(SanitizeFileName(sectionName));
        var pages = oneNote.GetPages(sectionId);

        foreach (var page in pages)
        {
            ExportPage(oneNote, converter, page.Id, page.Name, sectionDir);
        }
    }

    private static void ExportContainer(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter? converter,
        ResolvedRef container,
        DirectoryInfo outputDir)
    {
        var containerDir = outputDir.CreateSubdirectory(SanitizeFileName(container.Name));
        var sections = oneNote.GetSections(container.Id);

        foreach (var section in sections)
        {
            var sectionPath = section.SectionGroup is not null
                ? Path.Combine(containerDir.FullName, SanitizeFileName(section.SectionGroup), SanitizeFileName(section.Name))
                : Path.Combine(containerDir.FullName, SanitizeFileName(section.Name));

            var sectionDir = new DirectoryInfo(sectionPath);
            if (!sectionDir.Exists)
                sectionDir.Create();

            var pages = oneNote.GetPages(section.Id);
            foreach (var page in pages)
            {
                ExportPage(oneNote, converter, page.Id, page.Name, sectionDir);
            }
        }
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sanitized = new StringBuilder(name.Length);
        foreach (var c in name)
        {
            sanitized.Append(invalid.Contains(c) ? '_' : c);
        }
        return sanitized.ToString();
    }
}
