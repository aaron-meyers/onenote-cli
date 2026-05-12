using BlueMarsh.OneNote.CommandLine.OneNote;
using var oneNote = new OneNoteApplication();
var notebooks = oneNote.GetNotebooks();
var nb = notebooks.First(n => n.Name == "Personal");
var sections = oneNote.GetSections(nb.Id);
var section = sections.First(s => s.Name == "Gaming");
var xml = oneNote.GetHierarchy(section.Id, HierarchyScope.Pages);
Console.WriteLine(xml[..Math.Min(xml.Length, 2000)]);
