using System.CommandLine;
using System.Text;
using BlueMarsh.OneNote.CommandLine.Commands;

Console.InputEncoding = Encoding.UTF8;
Console.OutputEncoding = Encoding.UTF8;

var notebookCommand = new Command("notebook", "Manage notebooks")
{
    NotebookListCommand.Create(),
};

var sectionCommand = new Command("section", "Manage sections")
{
    SectionListCommand.Create(),
};

var pageCommand = new Command("page", "Manage pages")
{
    PageListCommand.Create(),
};

var rootCommand = new RootCommand("CLI for Microsoft OneNote")
{
    notebookCommand,
    sectionCommand,
    pageCommand,
};

var parseResult = rootCommand.Parse(args);
return await parseResult.InvokeAsync();
