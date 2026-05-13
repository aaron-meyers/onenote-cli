using System.CommandLine;
using System.Text;
using System.Xml.Linq;
using BlueMarsh.OneNote.CommandLine.Export;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

/// <summary>
/// Controls how the page title is represented in the exported Markdown.
/// </summary>
internal enum TitleMode
{
    /// <summary>
    /// Write the title as a YAML frontmatter property only if the filename
    /// does not match the title (e.g. when invalid filename characters were sanitized).
    /// Content headings are not offset.
    /// </summary>
    Auto,

    /// <summary>
    /// Do not include the title in the output at all.
    /// </summary>
    None,

    /// <summary>
    /// Always write the title as a YAML frontmatter property.
    /// Content headings are not offset.
    /// </summary>
    Property,

    /// <summary>
    /// Render the title as a Markdown # heading.
    /// Content headings are offset by +1 level.
    /// </summary>
    Heading,
}

/// <summary>
/// Controls which date properties are included in YAML frontmatter.
/// </summary>
internal enum DatesMode
{
    Both,
    Created,
    Updated,
    None,
}

internal static class PageExportCommand
{
    public static Command Create()
    {
        var refArg = new Argument<string?>("ref")
        {
            Description = "Page, section, section group, or notebook (name, ID, or path)",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var currentOption = new Option<bool>("--current", "-c")
        {
            Description = "Export pages currently open in OneNote",
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
        var titleOption = new Option<TitleMode>("--title", "-t")
        {
            Description = "How to handle the page title: auto (default), none, property, or heading",
            DefaultValueFactory = _ => TitleMode.Auto,
        };
        var datesOption = new Option<DatesMode>("--dates", "-d")
        {
            Description = "Which dates to include in frontmatter: both (default), created, updated, or none",
            DefaultValueFactory = _ => DatesMode.Both,
        };

        var command = new Command("export", "Export pages to Markdown or XML files")
        {
            refArg,
            currentOption,
            outputDirOption,
            rawOption,
            titleOption,
            datesOption,
        };

        command.SetAction(parseResult =>
        {
            var refString = parseResult.GetValue(refArg);
            var current = parseResult.GetValue(currentOption);
            var outputDir = parseResult.GetValue(outputDirOption)
                ?? new DirectoryInfo(Directory.GetCurrentDirectory());
            var raw = parseResult.GetValue(rawOption);
            var titleMode = parseResult.GetValue(titleOption);
            var datesMode = parseResult.GetValue(datesOption);

            if (current && refString is not null)
            {
                Console.Error.WriteLine("Cannot use --current with a ref argument.");
                return;
            }

            if (!current && refString is null)
            {
                Console.Error.WriteLine("Either a ref argument or --current must be specified.");
                return;
            }

            if (!outputDir.Exists)
                outputDir.Create();

            using var oneNote = new OneNoteApplication();

            OneNotePageToMarkdownConverter? converter = null;
            if (!raw)
            {
                converter = new OneNotePageToMarkdownConverter(
                    warn: msg => Console.Error.WriteLine($"Warning: {msg}"));
            }

            if (current)
            {
                var pages = oneNote.GetCurrentPages();
                if (pages.Count == 0)
                {
                    Console.Error.WriteLine("No pages are currently open in OneNote.");
                    return;
                }

                foreach (var page in pages)
                {
                    ExportPage(oneNote, converter, titleMode, datesMode, page.Id, page.Name, outputDir);
                }

                return;
            }

            var resolved = OneNoteRef.Resolve(oneNote, refString!);
            if (resolved is null)
            {
                Console.Error.WriteLine($"'{refString}' not found.");
                return;
            }

            switch (resolved.NodeType)
            {
                case HierarchyNodeType.Page:
                    ExportSinglePage(oneNote, converter, titleMode, datesMode, resolved, outputDir);
                    break;

                case HierarchyNodeType.Section:
                    ExportSection(oneNote, converter, titleMode, datesMode, resolved.Id, resolved.Name, outputDir);
                    break;

                case HierarchyNodeType.SectionGroup:
                case HierarchyNodeType.Notebook:
                    ExportContainer(oneNote, converter, titleMode, datesMode, resolved, outputDir);
                    break;
            }
        });

        return command;
    }

    private static void ExportPage(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter? converter,
        TitleMode titleMode,
        DatesMode datesMode,
        string pageId,
        string pageName,
        DirectoryInfo outputDir)
    {
        var pageXml = oneNote.GetPageContent(pageId);
        var sanitizedName = SanitizeFileName(pageName);

        string content;
        string extension;
        if (converter is not null)
        {
            var settings = new MarkdownConversionSettings
            {
                IncludeTitleProperty = titleMode switch
                {
                    TitleMode.Property => true,
                    TitleMode.Auto => sanitizedName != pageName,
                    _ => false,
                },
                IncludeTitleHeading = titleMode == TitleMode.Heading,
                IncludeCreatedDate = datesMode is DatesMode.Both or DatesMode.Created,
                IncludeUpdatedDate = datesMode is DatesMode.Both or DatesMode.Updated,
            };

            content = converter.Convert(pageXml, settings);
            extension = ".md";
        }
        else
        {
            content = XDocument.Parse(pageXml).ToString();
            extension = ".xml";
        }

        var filePath = Path.Combine(outputDir.FullName, sanitizedName + extension);
        File.WriteAllText(filePath, content, Encoding.UTF8);
        Console.WriteLine(filePath);
    }

    private static void ExportSinglePage(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter? converter,
        TitleMode titleMode,
        DatesMode datesMode,
        ResolvedRef pageRef,
        DirectoryInfo outputDir)
    {
        ExportPage(oneNote, converter, titleMode, datesMode, pageRef.Id, pageRef.Name, outputDir);
    }

    private static void ExportSection(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter? converter,
        TitleMode titleMode,
        DatesMode datesMode,
        string sectionId,
        string sectionName,
        DirectoryInfo outputDir)
    {
        var sectionDir = outputDir.CreateSubdirectory(SanitizeFileName(sectionName));
        var pages = oneNote.GetPages(sectionId);

        foreach (var page in pages)
        {
            ExportPage(oneNote, converter, titleMode, datesMode, page.Id, page.Name, sectionDir);
        }
    }

    private static void ExportContainer(
        OneNoteApplication oneNote,
        OneNotePageToMarkdownConverter? converter,
        TitleMode titleMode,
        DatesMode datesMode,
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
                ExportPage(oneNote, converter, titleMode, datesMode, page.Id, page.Name, sectionDir);
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
