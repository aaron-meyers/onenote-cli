using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace BlueMarsh.OneNote.CommandLine.Export;

/// <summary>
/// Converts OneNote page XML to Markdown using a recursive visitor pattern.
/// </summary>
internal sealed partial class OneNotePageToMarkdownConverter
{
    private static readonly XNamespace OneNoteNs =
        "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private readonly StringBuilder _output = new();
    private readonly Dictionary<int, QuickStyleInfo> _quickStyles = [];
    private readonly Dictionary<int, TagDefInfo> _tagDefs = [];
    private readonly Action<string> _warn;

    // Known elements that we handle or intentionally skip
    private static readonly HashSet<string> HandledElements =
    [
        "Page", "QuickStyleDef", "TagDef", "PageSettings", "PageSize",
        "Automatic", "RuleLines", "Title", "Outline", "Position", "Size",
        "OEChildren", "OE", "T", "List", "Bullet", "Number", "Tag",
        "Table", "Columns", "Column", "Row", "Cell",
        "Indents", "Indent", "Image", "CallbackID", "Data",
        "InsertedFile", "MediaFile", "InkDrawing", "XPSFile",
    ];

    public OneNotePageToMarkdownConverter(Action<string> warn)
    {
        _warn = warn;
    }

    public string Convert(string pageXml)
    {
        _output.Clear();
        _quickStyles.Clear();
        _tagDefs.Clear();

        var doc = XDocument.Parse(pageXml);
        var page = doc.Root;
        if (page is null)
            return "";

        VisitPage(page);

        return _output.ToString().TrimEnd() + "\n";
    }

    private void VisitPage(XElement page)
    {
        // Collect style and tag definitions
        foreach (var styleDef in page.Elements(OneNoteNs + "QuickStyleDef"))
        {
            var index = int.Parse(styleDef.Attribute("index")?.Value ?? "0");
            var name = styleDef.Attribute("name")?.Value ?? "";
            _quickStyles[index] = new QuickStyleInfo(index, name);
        }

        foreach (var tagDef in page.Elements(OneNoteNs + "TagDef"))
        {
            var index = int.Parse(tagDef.Attribute("index")?.Value ?? "0");
            var type = tagDef.Attribute("type")?.Value ?? "";
            var symbol = tagDef.Attribute("symbol")?.Value ?? "";
            _tagDefs[index] = new TagDefInfo(index, type, symbol);
        }

        // Title
        var title = page.Element(OneNoteNs + "Title");
        if (title is not null)
        {
            var titleText = ExtractTextFromOE(title.Element(OneNoteNs + "OE"));
            if (!string.IsNullOrWhiteSpace(titleText))
            {
                _output.AppendLine($"# {titleText}");
                _output.AppendLine();
            }
        }

        // Outlines
        foreach (var outline in page.Elements(OneNoteNs + "Outline"))
        {
            VisitOutline(outline);
        }

        // Warn about unhandled top-level elements
        foreach (var child in page.Elements())
        {
            var localName = child.Name.LocalName;
            if (!HandledElements.Contains(localName))
            {
                _warn($"Unhandled page-level element: <{localName}>");
            }
        }
    }

    private void VisitOutline(XElement outline)
    {
        var oeChildren = outline.Element(OneNoteNs + "OEChildren");
        if (oeChildren is not null)
        {
            VisitOEChildren(oeChildren, depth: 0);
        }
    }

    private void VisitOEChildren(XElement oeChildren, int depth)
    {
        foreach (var child in oeChildren.Elements())
        {
            if (child.Name.LocalName == "OE")
            {
                VisitOE(child, depth);
            }
            else if (!HandledElements.Contains(child.Name.LocalName))
            {
                _warn($"Unhandled element in OEChildren: <{child.Name.LocalName}>");
            }
        }
    }

    private void VisitOE(XElement oe, int depth)
    {
        // Check for table
        var table = oe.Element(OneNoteNs + "Table");
        if (table is not null)
        {
            VisitTable(table);
            return;
        }

        // Check for image (skip content but don't warn)
        if (oe.Element(OneNoteNs + "Image") is not null ||
            oe.Element(OneNoteNs + "InsertedFile") is not null ||
            oe.Element(OneNoteNs + "MediaFile") is not null ||
            oe.Element(OneNoteNs + "InkDrawing") is not null)
        {
            // Binary content — skip silently
            var children = oe.Element(OneNoteNs + "OEChildren");
            if (children is not null)
                VisitOEChildren(children, depth + 1);
            return;
        }

        // Determine line prefix based on list type, tag, and heading
        var prefix = "";
        var suffix = "";
        var isTodo = false;
        var isListItem = false;
        var isNumbered = false;

        // Check for tag (To-Do)
        var tag = oe.Element(OneNoteNs + "Tag");
        if (tag is not null)
        {
            var tagIndex = int.Parse(tag.Attribute("index")?.Value ?? "-1");
            var completed = tag.Attribute("completed")?.Value == "true";
            if (_tagDefs.TryGetValue(tagIndex, out var tagDef) && IsToDoTag(tagDef))
            {
                isTodo = true;
                prefix = completed ? "- [x] " : "- [ ] ";
            }
        }

        // Check for list
        var list = oe.Element(OneNoteNs + "List");
        if (list is not null && !isTodo)
        {
            if (list.Element(OneNoteNs + "Number") is not null)
            {
                isNumbered = true;
                prefix = "1. ";
            }
            else
            {
                prefix = "- ";
            }
            isListItem = true;
        }

        // Check for heading style
        var quickStyleIndex = int.Parse(oe.Attribute("quickStyleIndex")?.Value ?? "-1");
        var headingLevel = 0;
        if (_quickStyles.TryGetValue(quickStyleIndex, out var style))
        {
            headingLevel = style.Name switch
            {
                "h1" => 1,
                "h2" => 2,
                "h3" => 3,
                "h4" => 4,
                "h5" => 5,
                "h6" => 6,
                _ => 0,
            };
        }

        // Extract text
        var text = ExtractTextFromOE(oe);

        // Build the line
        if (headingLevel > 0 && !string.IsNullOrWhiteSpace(text))
        {
            // Add heading - title is h1 so content headings get +1
            var hashes = new string('#', headingLevel + 1);
            _output.AppendLine();
            _output.AppendLine($"{hashes} {text}");
            _output.AppendLine();
        }
        else if (isTodo || isListItem)
        {
            var indent = new string(' ', depth * 2);
            if (!string.IsNullOrEmpty(text) || isTodo)
            {
                _output.AppendLine($"{indent}{prefix}{text}");
            }
        }
        else
        {
            // Plain paragraph
            _output.AppendLine(text);
        }

        // Recurse into child OE elements
        var oeChildrenElement = oe.Element(OneNoteNs + "OEChildren");
        if (oeChildrenElement is not null)
        {
            var childDepth = (isTodo || isListItem) ? depth + 1 : depth;
            VisitOEChildren(oeChildrenElement, childDepth);
        }

        // Warn about unhandled child elements
        foreach (var child in oe.Elements())
        {
            var localName = child.Name.LocalName;
            if (!HandledElements.Contains(localName))
            {
                _warn($"Unhandled element in OE: <{localName}>");
            }
        }
    }

    private void VisitTable(XElement table)
    {
        var rows = table.Elements(OneNoteNs + "Row").ToList();
        if (rows.Count == 0)
            return;

        _output.AppendLine();

        // Process each row
        var isFirstRow = true;
        var columnCount = 0;

        foreach (var row in rows)
        {
            var cells = row.Elements(OneNoteNs + "Cell").ToList();
            columnCount = Math.Max(columnCount, cells.Count);

            var cellTexts = cells.Select(cell =>
            {
                var oeChildren = cell.Element(OneNoteNs + "OEChildren");
                if (oeChildren is null) return "";

                var parts = new List<string>();
                foreach (var oe in oeChildren.Elements(OneNoteNs + "OE"))
                {
                    var text = ExtractTextFromOE(oe);
                    if (!string.IsNullOrEmpty(text))
                        parts.Add(text);
                }
                // Join multi-line cell content with <br>
                return string.Join("<br>", parts);
            }).ToList();

            _output.Append('|');
            foreach (var text in cellTexts)
            {
                _output.Append($" {text.Replace("|", "\\|")} |");
            }
            _output.AppendLine();

            // After the first row, add the separator
            if (isFirstRow)
            {
                _output.Append('|');
                for (int i = 0; i < columnCount; i++)
                    _output.Append(" --- |");
                _output.AppendLine();
                isFirstRow = false;
            }
        }

        _output.AppendLine();
    }

    private string ExtractTextFromOE(XElement? oe)
    {
        if (oe is null) return "";

        var parts = new List<string>();
        foreach (var t in oe.Elements(OneNoteNs + "T"))
        {
            var cdata = t.Value;
            if (!string.IsNullOrEmpty(cdata))
            {
                parts.Add(ConvertInlineHtml(cdata));
            }
        }

        return string.Join("", parts).Trim();
    }

    private static string ConvertInlineHtml(string html)
    {
        if (!html.Contains('<'))
            return html;

        var result = new StringBuilder();
        var pos = 0;

        while (pos < html.Length)
        {
            var tagStart = html.IndexOf('<', pos);
            if (tagStart < 0)
            {
                result.Append(html, pos, html.Length - pos);
                break;
            }

            // Append text before the tag
            if (tagStart > pos)
                result.Append(html, pos, tagStart - pos);

            var tagEnd = html.IndexOf('>', tagStart);
            if (tagEnd < 0)
            {
                result.Append(html, pos, html.Length - pos);
                break;
            }

            var tag = html.AsSpan(tagStart, tagEnd - tagStart + 1);

            if (tag.StartsWith("<span".AsSpan(), StringComparison.OrdinalIgnoreCase))
            {
                var style = ExtractAttribute(tag, "style");
                var (mdOpen, mdClose) = ParseStyleToMarkdown(style);
                var spanEnd = html.IndexOf("</span>", tagEnd, StringComparison.OrdinalIgnoreCase);
                if (spanEnd >= 0)
                {
                    var innerHtml = html[(tagEnd + 1)..spanEnd];
                    var innerText = ConvertInlineHtml(innerHtml);
                    if (!string.IsNullOrWhiteSpace(innerText))
                    {
                        result.Append(mdOpen);
                        result.Append(innerText);
                        result.Append(mdClose);
                    }
                    pos = spanEnd + "</span>".Length;
                    continue;
                }
            }
            else if (tag.StartsWith("<br".AsSpan(), StringComparison.OrdinalIgnoreCase))
            {
                result.Append("  \n");
                pos = tagEnd + 1;
                continue;
            }
            else if (tag.StartsWith("<a ".AsSpan(), StringComparison.OrdinalIgnoreCase))
            {
                var href = ExtractAttribute(tag, "href");
                var anchorEnd = html.IndexOf("</a>", tagEnd, StringComparison.OrdinalIgnoreCase);
                if (anchorEnd >= 0)
                {
                    var linkText = ConvertInlineHtml(html[(tagEnd + 1)..anchorEnd]);
                    var hrefStr = href.ToString();
                    if (!string.IsNullOrEmpty(hrefStr))
                        result.Append($"[{linkText}]({hrefStr})");
                    else
                        result.Append(linkText);
                    pos = anchorEnd + "</a>".Length;
                    continue;
                }
            }

            // Skip unrecognized tags
            pos = tagEnd + 1;
        }

        return result.ToString();
    }

    private static ReadOnlySpan<char> ExtractAttribute(ReadOnlySpan<char> tag, string attrName)
    {
        var search = $"{attrName}=".AsSpan();
        var idx = tag.IndexOf(search, StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
            return [];

        idx += search.Length;
        if (idx >= tag.Length)
            return [];

        var quote = tag[idx];
        if (quote is '\'' or '"')
        {
            idx++;
            var end = tag[idx..].IndexOf(quote);
            if (end >= 0)
                return tag.Slice(idx, end);
        }

        return [];
    }

    private static (string open, string close) ParseStyleToMarkdown(ReadOnlySpan<char> style)
    {
        if (style.IsEmpty)
            return ("", "");

        var styleStr = style.ToString().ToLowerInvariant();
        var open = new StringBuilder();
        var close = new StringBuilder();

        if (styleStr.Contains("font-weight:bold") || styleStr.Contains("font-weight: bold"))
        {
            open.Append("**");
            close.Insert(0, "**");
        }

        if (styleStr.Contains("font-style:italic") || styleStr.Contains("font-style: italic"))
        {
            open.Append('*');
            close.Insert(0, '*');
        }

        if (styleStr.Contains("text-decoration:line-through") || styleStr.Contains("text-decoration: line-through"))
        {
            open.Append("~~");
            close.Insert(0, "~~");
        }

        return (open.ToString(), close.ToString());
    }

    private static bool IsToDoTag(TagDefInfo tagDef)
    {
        // To-Do tag: type="0" symbol="3" or type="0" symbol="0" (common patterns)
        // Also check for type="0" which is the standard checkbox
        return tagDef.Type == "0";
    }

    private sealed record QuickStyleInfo(int Index, string Name);
    private sealed record TagDefInfo(int Index, string Type, string Symbol);
}
