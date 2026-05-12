using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

internal static class NotebookListCommand
{
    public static Command Create()
    {
        var idOption = new Option<bool>("--id")
        {
            Description = "Show object IDs",
        };
        var command = new Command("list", "List all notebooks")
        {
            idOption,
        };

        command.SetAction(parseResult =>
        {
            var showId = parseResult.GetValue(idOption);

            using var oneNote = new OneNoteApplication();
            var notebooks = oneNote.GetNotebooks();

            if (notebooks.Count == 0)
            {
                Console.WriteLine("No notebooks found.");
                return;
            }

            foreach (var notebook in notebooks)
            {
                if (showId)
                    Console.WriteLine($"{notebook.Name} {notebook.Id}");
                else
                    Console.WriteLine(notebook.Name);
            }
        });

        return command;
    }
}
