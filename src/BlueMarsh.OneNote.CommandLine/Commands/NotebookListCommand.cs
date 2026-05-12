using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

internal static class NotebookListCommand
{
    public static Command Create()
    {
        var command = new Command("list", "List all notebooks");

        command.SetAction(parseResult =>
        {
            using var oneNote = new OneNoteApplication();
            var notebooks = oneNote.GetNotebooks();

            if (notebooks.Count == 0)
            {
                Console.WriteLine("No notebooks found.");
                return;
            }

            foreach (var notebook in notebooks)
            {
                Console.WriteLine(notebook.Name);
            }
        });

        return command;
    }
}
