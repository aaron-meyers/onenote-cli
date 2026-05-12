using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

internal static class SectionListCommand
{
    public static Command Create()
    {
        var containerArg = new Argument<string?>("container-ref")
        {
            Description = "Notebook or section group (name, ID, or path). Lists all if omitted",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var command = new Command("list", "List sections in a container");
        command.Add(containerArg);

        command.SetAction(parseResult =>
        {
            var containerRef = parseResult.GetValue(containerArg);

            using var oneNote = new OneNoteApplication();

            if (containerRef is null)
            {
                ListAllSections(oneNote);
            }
            else
            {
                var resolved = OneNoteRef.ResolveContainer(oneNote, containerRef);
                if (resolved is null)
                {
                    Console.Error.WriteLine($"Container '{containerRef}' not found. Must be a notebook or section group.");
                    return;
                }

                ListSectionsInContainer(oneNote, resolved);
            }
        });

        return command;
    }

    private static void ListAllSections(OneNoteApplication oneNote)
    {
        foreach (var notebook in oneNote.GetNotebooks())
        {
            var sections = oneNote.GetSections(notebook.Id);
            foreach (var section in sections)
            {
                var path = section.SectionGroup is not null
                    ? $"{notebook.Name}/{section.SectionGroup}/{section.Name}"
                    : $"{notebook.Name}/{section.Name}";
                Console.WriteLine(path);
            }
        }
    }

    private static void ListSectionsInContainer(OneNoteApplication oneNote, ResolvedRef container)
    {
        var sections = oneNote.GetSections(container.Id);
        foreach (var section in sections)
        {
            if (section.SectionGroup is not null)
                Console.WriteLine($"{section.SectionGroup}/{section.Name}");
            else
                Console.WriteLine(section.Name);
        }
    }
}
