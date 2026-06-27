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
onenote page export <ref> [--output-dir <path>] [--raw] [--line-endings LF|CRLF|System]
```

The `<ref>` argument accepts a notebook name, a `/`-separated path (e.g. `MyNotebook/Section/Page`), or a OneNote object ID.

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
- [OneNote to Markdown Exporter](https://github.com/segunak/one-note-to-markdown) - just discovered this as I was adding tags to my repo. Looks like it takes largely the same approach with the same goals as my project. I'll go ahead and archive mine if I end up switching to this instead.

## License

This project is licensed under the [MIT License](LICENSE).
