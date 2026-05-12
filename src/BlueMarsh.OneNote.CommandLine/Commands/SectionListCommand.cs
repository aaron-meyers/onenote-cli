using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

internal static class SectionListCommand
{
    public static Command Create()
    {
        var notebookArg = new Argument<string>("notebook") { Description = "Name of the notebook" };
        var command = new Command("list", "List sections in a notebook");
        command.Add(notebookArg);

        command.SetAction(parseResult =>
        {
            var notebookName = parseResult.GetValue(notebookArg)!;

            using var oneNote = new OneNoteApplication();
            var notebook = oneNote.FindNotebook(notebookName);

            if (notebook is null)
            {
                Console.Error.WriteLine($"Notebook '{notebookName}' not found.");
                return;
            }

            var sections = oneNote.GetSections(notebook.Id);

            if (sections.Count == 0)
            {
                Console.WriteLine("No sections found.");
                return;
            }

            foreach (var section in sections)
            {
                if (section.SectionGroup is not null)
                    Console.WriteLine($"{section.SectionGroup}/{section.Name}");
                else
                    Console.WriteLine(section.Name);
            }
        });

        return command;
    }
}
