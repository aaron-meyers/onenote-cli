namespace BlueMarsh.OneNote.CommandLine.OneNote;

/// <summary>
/// Specifies the scope of the hierarchy tree to retrieve from OneNote.
/// </summary>
internal enum HierarchyScope
{
    Self = 0,
    Children = 1,
    Notebooks = 2,
    Sections = 3,
    Pages = 4,
}

internal enum XMLSchema
{
    xs2007 = 0,
    xs2010 = 1,
    xs2013 = 2,
    xs2026 = 3,
}

internal enum CreateFileType
{
    cftNone = 0,
    cftNotebook = 1,
    cftFolder = 2,
    cftSection = 3,
}

internal enum NewPageStyle
{
    npsDefault = 0,
    npsBlankPageWithTitle = 1,
    npsBlankPageNoTitle = 2,
}

internal enum PageInfo
{
    piBasic = 0,
    piBinaryData = 1,
    piSelection = 2,
    piBinaryDataSelection = 3,
    piFileType = 4,
    piBinaryDataFileType = 5,
    piSelectionFileType = 6,
    piAll = 7,
}

internal enum PublishFormat
{
    pfOneNote = 0,
    pfOneNotePackage = 1,
    pfMHTML = 2,
    pfPDF = 3,
    pfXPS = 4,
    pfWord = 5,
    pfEMF = 6,
    pfHTML = 7,
    pfOneNote2007 = 8,
}

internal enum SpecialLocation
{
    slBackUpFolder = 0,
    slUnfiledNotesSection = 1,
    slDefaultNotebookFolder = 2,
}

internal enum FilingLocation
{
    flEMail = 0,
    flContacts = 1,
    flTasks = 2,
    flMeetings = 3,
    flWebContent = 4,
    flPrintOuts = 5,
}

internal enum FilingLocationType
{
    fltNamedSectionNewPage = 0,
    fltCurrentSectionNewPage = 1,
    fltCurrentPage = 2,
    fltNamedPage = 3,
}
