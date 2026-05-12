using System.CommandLine;
using System.Text;
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
            Description = "Output directory for exported Markdown files (defaults to current directory)",
        };

        var command = new Command("export", "Export pages to Markdown files");
        command.Add(refArg);
        command.Add(outputDirOption);

        command.SetAction(parseResult =>
        {
            var refString = parseResult.GetValue(refArg)!;
            var outputDir = parseResult.GetValue(outputDirOption)
                ?? new DirectoryInfo(Directory.GetCurrentDirectory());

            if (!outputDir.Exists)
                outputDir.Create();

            using var oneNote = new OneNoteApplication();

            var resolved = OneNoteRef.Resolve(oneNote, refString);
            if (resolved is null)
            {
                Console.Error.WriteLine($"'{refString}' not found.");
                return;
            }

            var converter = new OneNotePageToMarkdownConverter(
                warn: msg => Console.Error.WriteLine($"Warning: {msg}"));

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

    private static void ExportSinglePage(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter converter,
        ResolvedRef pageRef,
        DirectoryInfo outputDir)
    {
        var pageXml = oneNote.GetPageContent(pageRef.Id);
        var markdown = converter.Convert(pageXml);
        var filePath = Path.Combine(outputDir.FullName, SanitizeFileName(pageRef.Name) + ".md");
        File.WriteAllText(filePath, markdown, Encoding.UTF8);
        Console.WriteLine(filePath);
    }

    private static void ExportSection(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter converter,
        string sectionId,
        string sectionName,
        DirectoryInfo outputDir)
    {
        var sectionDir = outputDir.CreateSubdirectory(SanitizeFileName(sectionName));
        var pages = oneNote.GetPages(sectionId);

        foreach (var page in pages)
        {
            var pageXml = oneNote.GetPageContent(page.Id);
            var markdown = converter.Convert(pageXml);
            var filePath = Path.Combine(sectionDir.FullName, SanitizeFileName(page.Name) + ".md");
            File.WriteAllText(filePath, markdown, Encoding.UTF8);
            Console.WriteLine(filePath);
        }
    }

    private static void ExportContainer(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter converter,
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
                var pageXml = oneNote.GetPageContent(page.Id);
                var markdown = converter.Convert(pageXml);
                var filePath = Path.Combine(sectionDir.FullName, SanitizeFileName(page.Name) + ".md");
                File.WriteAllText(filePath, markdown, Encoding.UTF8);
                Console.WriteLine(filePath);
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
