# onenote-cli

A command-line interface for Microsoft OneNote on Windows. Exports notebooks, sections, and pages to Markdown or XML by communicating with the desktop OneNote application via COM interop.

## Requirements

- Windows with Microsoft OneNote desktop application installed
- .NET 10 SDK

## Usage

```
onenote notebook list
onenote section list <notebook>
onenote page list <notebook/section>
onenote page export <ref> [--output-dir <path>] [--raw]
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

## License

This project is licensed under the [MIT License](LICENSE).
