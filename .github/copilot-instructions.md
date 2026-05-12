# Copilot Instructions

## Build & Test

```powershell
# Build
dotnet build

# Run
dotnet run --project src/BlueMarsh.OneNote.CommandLine -- <command>

# Run all tests
dotnet test

# Run a single test by name
dotnet test --filter "MethodName"
```

Tests use **TUnit** (not xUnit/NUnit) with **Verify** for snapshot testing. Snapshot files live alongside the test file as `*.verified.txt`. When converter output changes, update snapshots with `dotnet test` (DiffEngine is configured with `AutoLaunch: false`).

## Architecture

This is a .NET 10 Windows CLI that interacts with desktop OneNote via COM interop. It targets `net10.0-windows`.

- **OneNote COM layer** (`src/.../OneNote/`): `OneNoteApplication` wraps the COM `IOneNoteApplication` interface using `[ComImport]` vtable binding (not IDispatch) because the type library is only registered for Win32. It exposes hierarchy queries (notebooks → sections → pages) and page content as XML.
- **Reference resolution** (`OneNoteRef`): Resolves user-provided strings (names, IDs containing `{`, or `/`-separated paths like `Notebook/Section/Page`) to hierarchy objects.
- **Export** (`src/.../Export/`): `OneNotePageToMarkdownConverter` converts OneNote page XML to Markdown using a recursive visitor. It handles headings, lists, to-dos, code blocks, tables, inline formatting, and links.
- **Commands** (`src/.../Commands/`): Each command is a static class with a `Create()` method returning a `System.CommandLine.Command`. Commands are composed into a hierarchy in `Program.cs`.

## Conventions

- **System.CommandLine 2.0.7 API**: Use `Add()` (not `AddCommand()`). `Option` aliases go in the constructor. Use `Parse(args).InvokeAsync()` for invocation.
- **Command pattern**: Commands are static classes with a `Create()` method. Action handlers use `command.SetAction(parseResult => { ... })`.
- **Records for data**: OneNote hierarchy items use `sealed record` types (`NotebookInfo`, `SectionInfo`, `PageInfo`, `ResolvedRef`).
- **Internal visibility**: All non-program types are `internal`. Tests access internals via `InternalsVisibleTo`.
- **Nullable enabled**: All code uses `<Nullable>enable</Nullable>`.
