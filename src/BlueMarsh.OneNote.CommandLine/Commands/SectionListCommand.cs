using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.OneNote;

namespace BlueMarsh.OneNote.CommandLine.Commands;

internal static class SectionListCommand
{
    public static Command Create()
    {
        var notebookArg = new Argument<string?>("notebook") { Description = "Name of the notebook (lists all if omitted)", Arity = ArgumentArity.ZeroOrOne };
        var command = new Command("list", "List sections in a notebook");
        command.Add(notebookArg);

        command.SetAction(parseResult =>
        {
            var notebookName = parseResult.GetValue(notebookArg);

            using var oneNote = new OneNoteApplication();
            var notebooks = oneNote.GetNotebooks();

            IEnumerable<NotebookInfo> targets;
            if (notebookName is not null)
            {
                var notebook = notebooks.FirstOrDefault(n => n.Name.Equals(notebookName, StringComparison.OrdinalIgnoreCase));
                if (notebook is null)
                {
                    Console.Error.WriteLine($"Notebook '{notebookName}' not found.");
                    return;
                }
                targets = [notebook];
            }
            else
            {
                targets = notebooks;
            }

            foreach (var notebook in targets)
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
        });

        return command;
    }
}
