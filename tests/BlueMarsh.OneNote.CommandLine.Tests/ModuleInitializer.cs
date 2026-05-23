using System.Runtime.CompilerServices;

namespace BlueMarsh.OneNote.CommandLine.Tests;

internal static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        DiffEngine.DiffRunner.Disabled = true;
    }
}
