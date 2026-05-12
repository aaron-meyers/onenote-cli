using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.Commands;

var notebookCommand = new Command("notebook", "Manage notebooks");
notebookCommand.Add(NotebookListCommand.Create());

var rootCommand = new RootCommand("CLI for Microsoft OneNote");
rootCommand.Add(notebookCommand);

var parseResult = rootCommand.Parse(args);
return await parseResult.InvokeAsync();
