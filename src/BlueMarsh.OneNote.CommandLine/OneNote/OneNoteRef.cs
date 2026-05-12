using System.Xml.Linq;

namespace BlueMarsh.OneNote.CommandLine.OneNote;

/// <summary>
/// Represents a resolved reference to a OneNote hierarchy object.
/// </summary>
internal sealed record ResolvedRef(
    string Id,
    string Name,
    HierarchyNodeType NodeType,
    string FullPath);

internal enum HierarchyNodeType
{
    Notebook,
    SectionGroup,
    Section,
    Page,
}

/// <summary>
/// Parses and resolves OneNote object references. A ref can be a name,
/// an ID (containing '{'), or a path with '/' separators.
/// </summary>
internal static class OneNoteRef
{
    private static readonly XNamespace OneNoteNs =
        "http://schemas.microsoft.com/office/onenote/2013/onenote";

    /// <summary>
    /// Resolves a ref string to a hierarchy object.
    /// Returns null if the ref cannot be resolved.
    /// </summary>
    public static ResolvedRef? Resolve(OneNoteApplication oneNote, string refString)
    {
        if (IsId(refString))
            return ResolveById(oneNote, refString);

        if (refString.Contains('/'))
            return ResolveByPath(oneNote, refString);

        return ResolveByName(oneNote, refString);
    }

    /// <summary>
    /// Resolves a ref that must be a container (notebook or section group).
    /// Returns null if the ref cannot be resolved or is not a container.
    /// </summary>
    public static ResolvedRef? ResolveContainer(OneNoteApplication oneNote, string refString)
    {
        var resolved = Resolve(oneNote, refString);
        if (resolved is null)
            return null;

        if (resolved.NodeType is not (HierarchyNodeType.Notebook or HierarchyNodeType.SectionGroup))
            return null;

        return resolved;
    }

    private static bool IsId(string refString) => refString.Contains('{');

    private static ResolvedRef? ResolveById(OneNoteApplication oneNote, string id)
    {
        var xml = oneNote.GetHierarchy(id, HierarchyScope.Self);
        var doc = XDocument.Parse(xml);
        var root = doc.Root;
        if (root is null)
            return null;

        var nodeType = root.Name.LocalName switch
        {
            "Notebook" or "Notebooks" => HierarchyNodeType.Notebook,
            "SectionGroup" => HierarchyNodeType.SectionGroup,
            "Section" => HierarchyNodeType.Section,
            "Page" => HierarchyNodeType.Page,
            _ => (HierarchyNodeType?)null,
        };

        if (nodeType is null)
            return null;

        var name = root.Attribute("name")?.Value ?? "";
        return new ResolvedRef(id, name, nodeType.Value, name);
    }

    private static ResolvedRef? ResolveByName(OneNoteApplication oneNote, string name)
    {
        var notebooks = oneNote.GetNotebooks();
        var notebook = notebooks.FirstOrDefault(
            n => n.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

        if (notebook is not null)
            return new ResolvedRef(notebook.Id, notebook.Name, HierarchyNodeType.Notebook, notebook.Name);

        return null;
    }

    private static ResolvedRef? ResolveByPath(OneNoteApplication oneNote, string path)
    {
        var segments = path.Split('/');
        if (segments.Length == 0)
            return null;

        // First segment is the notebook name
        var notebooks = oneNote.GetNotebooks();
        var notebook = notebooks.FirstOrDefault(
            n => n.Name.Equals(segments[0], StringComparison.OrdinalIgnoreCase));
        if (notebook is null)
            return null;

        if (segments.Length == 1)
            return new ResolvedRef(notebook.Id, notebook.Name, HierarchyNodeType.Notebook, notebook.Name);

        // Walk section groups
        var xml = oneNote.GetHierarchy(notebook.Id, HierarchyScope.Sections);
        var currentElement = XDocument.Parse(xml).Root;
        var pathSoFar = notebook.Name;

        for (int i = 1; i < segments.Length; i++)
        {
            if (currentElement is null)
                return null;

            var segment = segments[i];
            var sectionGroup = currentElement
                .Elements(OneNoteNs + "SectionGroup")
                .FirstOrDefault(e => (e.Attribute("name")?.Value ?? "")
                    .Equals(segment, StringComparison.OrdinalIgnoreCase));

            if (sectionGroup is not null)
            {
                pathSoFar = $"{pathSoFar}/{sectionGroup.Attribute("name")?.Value ?? segment}";
                currentElement = sectionGroup;
                continue;
            }

            // Check if it matches a section (not a container, but still a valid ref)
            var section = currentElement
                .Elements(OneNoteNs + "Section")
                .FirstOrDefault(e => (e.Attribute("name")?.Value ?? "")
                    .Equals(segment, StringComparison.OrdinalIgnoreCase));

            if (section is not null)
            {
                var sectionName = section.Attribute("name")?.Value ?? segment;
                var sectionId = section.Attribute("ID")?.Value ?? "";
                var sectionFullPath = $"{pathSoFar}/{sectionName}";

                // If this is the last segment, return the section
                if (i == segments.Length - 1)
                {
                    return new ResolvedRef(sectionId, sectionName, HierarchyNodeType.Section, sectionFullPath);
                }

                // Otherwise remaining segments must resolve to a page within this section
                return ResolvePageInSection(oneNote, sectionId, sectionFullPath, segments[(i + 1)..]);
            }

            return null;
        }

        // Ended on a section group
        var id = currentElement?.Attribute("ID")?.Value ?? "";
        var name = currentElement?.Attribute("name")?.Value ?? segments[^1];
        return new ResolvedRef(id, name, HierarchyNodeType.SectionGroup, pathSoFar);
    }

    private static ResolvedRef? ResolvePageInSection(
        OneNoteApplication oneNote, string sectionId, string sectionPath, ReadOnlySpan<string> pageSegments)
    {
        // Currently only single-segment page names are supported
        if (pageSegments.Length != 1)
            return null;

        var pageName = pageSegments[0];
        var xml = oneNote.GetHierarchy(sectionId, HierarchyScope.Pages);
        var doc = XDocument.Parse(xml);

        var page = doc.Root?
            .Elements(OneNoteNs + "Page")
            .FirstOrDefault(e => (e.Attribute("name")?.Value ?? "")
                .Equals(pageName, StringComparison.OrdinalIgnoreCase));

        if (page is null)
            return null;

        var pageActualName = page.Attribute("name")?.Value ?? pageName;
        return new ResolvedRef(
            page.Attribute("ID")?.Value ?? "",
            pageActualName,
            HierarchyNodeType.Page,
            $"{sectionPath}/{pageActualName}");
    }
}
