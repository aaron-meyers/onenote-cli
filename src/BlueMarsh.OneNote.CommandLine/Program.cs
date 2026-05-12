using System.CommandLine;
using System.Text;
using BlueMarsh.OneNote.CommandLine.Commands;

Console.InputEncoding = Encoding.UTF8;
Console.OutputEncoding = Encoding.UTF8;

var rootCommand = new RootCommand("CLI for Microsoft OneNote")
{
    new Command("notebook", "Manage notebooks")
    {
        NotebookListCommand.Create(),
    },
    new Command("section", "Manage sections")
    {
        SectionListCommand.Create(),
    },
    new Command("page", "Manage pages")
    {
        PageListCommand.Create(),
        PageExportCommand.Create(),
    },
};

var parseResult = rootCommand.Parse(args);
return await parseResult.InvokeAsync();






