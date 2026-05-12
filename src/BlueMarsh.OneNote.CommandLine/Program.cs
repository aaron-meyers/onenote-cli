using System.CommandLine;
using BlueMarsh.OneNote.CommandLine.Commands;

var notebookCommand = new Command("notebook", "Manage notebooks");
notebookCommand.Add(NotebookListCommand.Create());

var sectionCommand = new Command("section", "Manage sections");
sectionCommand.Add(SectionListCommand.Create());

var pageCommand = new Command("page", "Manage pages");
pageCommand.Add(PageListCommand.Create());

var rootCommand = new RootCommand("CLI for Microsoft OneNote");
rootCommand.Add(notebookCommand);
rootCommand.Add(sectionCommand);
rootCommand.Add(pageCommand);

var parseResult = rootCommand.Parse(args);
return await parseResult.InvokeAsync();
