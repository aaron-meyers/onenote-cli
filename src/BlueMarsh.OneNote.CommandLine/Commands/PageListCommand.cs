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
        var command = new Command("list", "List pages in a container")
        {
            containerArg,
        };

        command.SetAction(parseResult =>
        {
            var containerRef = parseResult.GetValue(containerArg);

            using var oneNote = new OneNoteApplication();

            if (containerRef is null)
            {
                ListAllPages(oneNote);
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

                ListPages(oneNote, resolved);
            }
        });

        return command;
    }

    private static void ListAllPages(OneNoteApplication oneNote)
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
                    Console.WriteLine($"{sectionPath}/{page.Name}");
                }
            }
        }
    }

    private static void ListPages(OneNoteApplication oneNote, ResolvedRef container)
    {
        if (container.NodeType == HierarchyNodeType.Section)
        {
            ListPagesInSection(oneNote, container.Id);
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
                    Console.WriteLine($"{sectionPath}/{page.Name}");
                }
            }
        }
    }

    private static void ListPagesInSection(OneNoteApplication oneNote, string sectionId)
    {
        var pages = oneNote.GetPages(sectionId);
        foreach (var page in pages)
        {
            var indent = new string(' ', (page.Level - 1) * 2);
            Console.WriteLine($"{indent}{page.Name}");
        }
    }
}
