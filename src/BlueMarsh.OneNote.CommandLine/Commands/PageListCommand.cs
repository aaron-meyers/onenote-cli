using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

internal static class PageListCommand
{
    public static Command Create()
    {
        var containerArg = new Argument<string?>("container-ref")
        {
            Description = "Notebook, section group, or section (name, ID, or path). Lists all if omitted",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var idOption = new Option<bool>("--id")
        {
            Description = "Show object IDs",
        };
        var command = new Command("list", "List pages in a container")
        {
            containerArg,
            idOption,
        };

        command.SetAction(parseResult =>
        {
            var containerRef = parseResult.GetValue(containerArg);
            var showId = parseResult.GetValue(idOption);

            using var oneNote = new OneNoteApplication();

            if (containerRef is null)
            {
                ListAllPages(oneNote, showId);
            }
            else
            {
                var resolved = OneNoteRef.Resolve(oneNote, containerRef);
                if (resolved is null)
                {
                    Console.Error.WriteLine($"'{containerRef}' not found.");
                    return;
                }

                if (resolved.NodeType == HierarchyNodeType.Page)
                {
                    Console.Error.WriteLine($"'{containerRef}' is a page, not a container.");
                    return;
                }

                ListPages(oneNote, resolved, showId);
            }
        });

        return command;
    }

    private static void ListAllPages(OneNoteApplication oneNote, bool showId)
    {
        foreach (var notebook in oneNote.GetNotebooks())
        {
            var sections = oneNote.GetSections(notebook.Id);
            foreach (var section in sections)
            {
                var sectionPath = section.SectionGroup is not null
                    ? $"{notebook.Name}/{section.SectionGroup}/{section.Name}"
                    : $"{notebook.Name}/{section.Name}";

                var pages = oneNote.GetPages(section.Id);
                foreach (var page in pages)
                {
                    if (showId)
                        Console.WriteLine($"{sectionPath}/{page.Name} {page.Id}");
                    else
                        Console.WriteLine($"{sectionPath}/{page.Name}");
                }
            }
        }
    }

    private static void ListPages(OneNoteApplication oneNote, ResolvedRef container, bool showId)
    {
        if (container.NodeType == HierarchyNodeType.Section)
        {
            ListPagesInSection(oneNote, container.Id, showId);
        }
        else
        {
            // Notebook or section group — list pages with section path prefix
            var sections = oneNote.GetSections(container.Id);
            foreach (var section in sections)
            {
                var sectionPath = section.SectionGroup is not null
                    ? $"{section.SectionGroup}/{section.Name}"
                    : section.Name;

                var pages = oneNote.GetPages(section.Id);
                foreach (var page in pages)
                {
                    if (showId)
                        Console.WriteLine($"{sectionPath}/{page.Name} {page.Id}");
                    else
                        Console.WriteLine($"{sectionPath}/{page.Name}");
                }
            }
        }
    }

    private static void ListPagesInSection(OneNoteApplication oneNote, string sectionId, bool showId)
    {
        var pages = oneNote.GetPages(sectionId);
        foreach (var page in pages)
        {
            var indent = new string(' ', (page.Level - 1) * 2);
            if (showId)
                Console.WriteLine($"{indent}{page.Name} {page.Id}");
            else
                Console.WriteLine($"{indent}{page.Name}");
        }
    }
}
