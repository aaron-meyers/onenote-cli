using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace BlueMarsh.OneNote.CommandLine.OneNote;

/// <summary>
/// Minimal COM interface for the OneNote Application object.
/// Uses vtable binding to avoid requiring a registered type library.
/// </summary>
[ComImport]
[Guid("452AC71A-B655-4967-A208-A4CC39DD7949")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
internal interface IOneNoteApplication
{
    void GetHierarchy(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrStartNodeID,
        [In] HierarchyScope hsScope,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrHierarchyXmlOut,
        [In, Optional, DefaultParameterValue(XMLSchema.xs2013)] XMLSchema xsSchema);

    void UpdateHierarchy(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrChangesXmlIn,
        [In, Optional, DefaultParameterValue(XMLSchema.xs2013)] XMLSchema xsSchema);

    void OpenHierarchy(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrPath,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrRelativeToObjectID,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrObjectID,
        [In] CreateFileType cftIfNotExist);

    void DeleteHierarchy(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        [In] DateTime dateExpectedLastModified,
        [In, Optional, DefaultParameterValue(false)] bool deletePermanently);

    void CreateNewPage(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrSectionID,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrPageID,
        [In, Optional, DefaultParameterValue(NewPageStyle.npsDefault)] NewPageStyle npsNewPageStyle);

    void CloseNotebook(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrNotebookID,
        [In, Optional, DefaultParameterValue(false)] bool force);

    void GetHierarchyParent(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrParentID);

    void GetPageContent(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrPageID,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrPageXmlOut,
        [In, Optional, DefaultParameterValue(PageInfo.piBasic)] PageInfo pageInfoToExport,
        [In, Optional, DefaultParameterValue(XMLSchema.xs2013)] XMLSchema xsSchema);

    void UpdatePageContent(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrPageChangesXmlIn,
        [In] DateTime dateExpectedLastModified,
        [In, Optional, DefaultParameterValue(XMLSchema.xs2013)] XMLSchema xsSchema,
        [In, Optional, DefaultParameterValue(false)] bool force);

    void GetBinaryPageContent(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrPageID,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrCallbackID,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrBinaryObjectB64Out);

    void DeletePageContent(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrPageID,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        [In] DateTime dateExpectedLastModified,
        [In, Optional, DefaultParameterValue(false)] bool force);

    void NavigateTo(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrHierarchyObjectID,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        [In, Optional, DefaultParameterValue(false)] bool fNewWindow);

    void NavigateToUrl(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrUrl,
        [In, Optional, DefaultParameterValue(false)] bool fNewWindow);

    void Publish(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrTargetFilePath,
        [In] PublishFormat pfPublishFormat,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrCLSIDofExporter);

    void OpenPackage(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrPathPackage,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrPathDest,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrPathOut);

    void GetHyperlinkToObject(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrPageContentObjectID,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrHyperlinkOut);

    void FindPages(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrStartNodeID,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrSearchString,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrHierarchyXmlOut,
        [In, Optional, DefaultParameterValue(false)] bool fIncludeUnindexedPages,
        [In, Optional, DefaultParameterValue(false)] bool fDisplay,
        [In, Optional, DefaultParameterValue(XMLSchema.xs2013)] XMLSchema xsSchema);

    void FindMeta(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrStartNodeID,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrSearchStringName,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrHierarchyXmlOut,
        [In, Optional, DefaultParameterValue(false)] bool fIncludeUnindexedPages,
        [In, Optional, DefaultParameterValue(XMLSchema.xs2013)] XMLSchema xsSchema);

    void GetSpecialLocation(
        [In] SpecialLocation slToGet,
        [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrSpecialLocationPath);

    void MergeFiles(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrBaseFile,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrClientFile,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrServerFile,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrTargetFile);

    [return: MarshalAs(UnmanagedType.IDispatch)]
    object QuickFiling();

    void SyncHierarchy(
        [In, MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID);

    void SetFilingLocation(
        [In] FilingLocation flToSet,
        [In] FilingLocationType fltToSet,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrFilingSectionID);
}

/// <summary>
/// Wraps the OneNote.Application COM object, providing managed access
/// to the OneNote interop API.
/// </summary>
internal sealed class OneNoteApplication : IDisposable
{
    private const string OneNoteProgId = "OneNote.Application";

    private static readonly XNamespace OneNoteNs =
        "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private readonly IOneNoteApplication _application;
    private bool _disposed;

    public OneNoteApplication()
    {
        var type = Type.GetTypeFromProgID(OneNoteProgId)
            ?? throw new InvalidOperationException(
                "OneNote is not installed or the COM class is not registered.");

        var instance = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException(
                "Failed to create an instance of the OneNote application.");

        _application = (IOneNoteApplication)instance;
    }

    /// <summary>
    /// Returns the hierarchy XML from OneNote at the given scope.
    /// </summary>
    public string GetHierarchy(string startNodeId, HierarchyScope scope)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _application.GetHierarchy(startNodeId, scope, out var xml);
        return xml;
    }

    /// <summary>
    /// Returns the list of notebook names and IDs from the hierarchy.
    /// </summary>
    public IReadOnlyList<NotebookInfo> GetNotebooks()
    {
        var xml = GetHierarchy("", HierarchyScope.Notebooks);
        var doc = XDocument.Parse(xml);

        return doc.Root?
            .Elements(OneNoteNs + "Notebook")
            .Select(e => new NotebookInfo(
                Id: e.Attribute("ID")?.Value ?? "",
                Name: e.Attribute("name")?.Value ?? "",
                NickName: e.Attribute("nickname")?.Value,
                Path: e.Attribute("path")?.Value ?? "",
                Color: e.Attribute("color")?.Value,
                LastModifiedTime: e.Attribute("lastModifiedTime")?.Value))
            .ToList()
            ?? [];
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        Marshal.ReleaseComObject(_application);
    }
}

internal sealed record NotebookInfo(
    string Id,
    string Name,
    string? NickName,
    string Path,
    string? Color,
    string? LastModifiedTime);
