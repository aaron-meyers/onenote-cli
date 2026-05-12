using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

internal static class SectionListCommand
{
    public static Command Create()
    {
        var containerArg = new Argument<string?>("container-ref")
        {
            Description = "Notebook or section group (name, ID, or path). Lists all if omitted.",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var idOption = new Option<bool>("--id")
        {
            Description = "Show object IDs",
        };
        var command = new Command("list", "List sections in a container")
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
                ListAllSections(oneNote, showId);
            }
            else
            {
                var resolved = OneNoteRef.ResolveContainer(oneNote, containerRef);
                if (resolved is null)
                {
                    Console.Error.WriteLine($"Container '{containerRef}' not found. Must be a notebook or section group.");
                    return;
                }

                ListSectionsInContainer(oneNote, resolved, showId);
            }
        });

        return command;
    }

    private static void ListAllSections(OneNoteApplication oneNote, bool showId)
    {
        foreach (var notebook in oneNote.GetNotebooks())
        {
            var sections = oneNote.GetSections(notebook.Id);
            foreach (var section in sections)
            {
                var path = section.SectionGroup is not null
                    ? $"{notebook.Name}/{section.SectionGroup}/{section.Name}"
                    : $"{notebook.Name}/{section.Name}";
                if (showId)
                    Console.WriteLine($"{path} {section.Id}");
                else
                    Console.WriteLine(path);
            }
        }
    }

    private static void ListSectionsInContainer(OneNoteApplication oneNote, ResolvedRef container, bool showId)
    {
        var sections = oneNote.GetSections(container.Id);
        foreach (var section in sections)
        {
            var name = section.SectionGroup is not null
                ? $"{section.SectionGroup}/{section.Name}"
                : section.Name;
            if (showId)
                Console.WriteLine($"{name} {section.Id}");
            else
                Console.WriteLine(name);
        }
    }
}
