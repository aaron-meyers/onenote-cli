# onenote-cli

A command-line interface for Microsoft OneNote on Windows. Exports notebooks, sections, and pages to Markdown or [MS-ONE](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-one/73d22548-a613-4350-8c23-07d15576be50) XML by communicating with the desktop OneNote application via [COM interop](https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote).

## Requirements

- Windows with Microsoft OneNote desktop application installed
- .NET 10 SDK

## Usage

```
onenote notebook list
onenote section list <notebook>
onenote page list <notebook/section>
onenote page export <ref> [--output-dir <path>] [--raw] ...
```

The `<ref>` argument accepts a notebook name, a `/`-separated path (e.g. `MyNotebook/Section/Page`), or a OneNote object ID.
The `export` command has additional options, `onenote page export --help` for more details.

### Examples

```powershell
# List all notebooks
onenote notebook list

# List sections in a notebook
onenote section list "My Notebook"

# Export a single page to Markdown
onenote page export "My Notebook/Notes/Meeting Notes" --output-dir ./export

# Export an entire notebook
onenote page export "My Notebook" -o ./export

# Export raw OneNote XML instead of Markdown
onenote page export "My Notebook/Notes/Meeting Notes" --raw
```

## Building

```powershell
dotnet build
```

To run directly from source:

```powershell
dotnet run --project src/BlueMarsh.OneNote.CommandLine -- <command>
```

## Testing

```powershell
# Run all tests
dotnet test

# Run a specific test
dotnet test --filter "PlainText"
```

Tests use [TUnit](https://github.com/thomhurst/TUnit) with [Verify](https://github.com/VerifyTests/Verify) for snapshot testing. Expected outputs are stored as `*.verified.txt` files alongside the test class.

## Related projects
- [onenote-cli](https://github.com/snomiao/onenote-cli) - Node.js CLI based on Microsoft Graph API for OneNote Online which similarly provides export of OneNote content to markdown. The Graph endpoints work with "OneNote-flavored HTML" rather than the [MS-ONE](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-one/73d22548-a613-4350-8c23-07d15576be50) XML format exposed by the Windows COM interface. My primary motivation for building a new project based on the COM interface is that I cannot access my enterprise notebooks via Microsoft Graph due to corporate IT restrictions.
- [OneNote to Markdown Exporter](https://github.com/segunak/one-note-to-markdown) - discovered this as I was adding tags to my repo. It takes the same core approach (OneNote COM interop) with the same high-level goal as my project. I am continuing with this project rather than contributing to that one because:
    - There are significant gaps in content conversion (e.g. broken tables, poor handling of adjacent spans with the same formatting, missing strikethrough/superscript/subscript handling, missing tag handling) that I've already addressed here. There are other gaps that I haven't addressed yet (as of writing) but intend to (e.g. missing attachments like PDFs). Unlike the author of the other project, I was a heavy user of OneNote and loved it for 20+ years. As a result, I have a significant amount of varied and valued content in OneNote so high fidelity conversion is a high priority for me.
    - It is also built as GUI-first and while it supports CLI commands, it doesn't function as a native CLI - the GUI temporarily appears during export and the CLI output doesn't work well in PowerShell (known PowerShell behavior when invoking GUI applications that write output to the console).
    - Finally, I have some different preferences for folder and file structure. I'm happy to add more control over this if others end up using my CLI but I'll initially focus on a layout that matches my new structure in Obsidian.

## License

This project is licensed under the [MIT License](LICENSE).
