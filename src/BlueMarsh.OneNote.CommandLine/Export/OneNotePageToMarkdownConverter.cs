using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace BlueMarsh.OneNote.CommandLine.Export;

/// <summary>
/// Settings controlling how a OneNote page is converted to Markdown.
/// </summary>
internal sealed record MarkdownConversionSettings
{
    /// <summary>If true, write the title as a YAML frontmatter property.</summary>
    public bool IncludeTitleProperty { get; init; }

    /// <summary>If true, render the title as a # heading and offset content headings by +1.</summary>
    public bool IncludeTitleHeading { get; init; }

    /// <summary>If true, include the created date in YAML frontmatter.</summary>
    public bool IncludeCreatedDate { get; init; }

    /// <summary>If true, include the updated date in YAML frontmatter.</summary>
    public bool IncludeUpdatedDate { get; init; }

    /// <summary>If true, write dates in UTC with a 'Z' suffix; otherwise use local timezone.</summary>
    public bool UtcDates { get; init; }
}

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
    private bool _includeTitleHeading;
    private bool _offsetHeadings;

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

    /// <summary>
    /// Converts OneNote page XML to Markdown.
    /// </summary>
    public string Convert(string pageXml, MarkdownConversionSettings? settings = null)
    {
        settings ??= new MarkdownConversionSettings();

        _output.Clear();
        _quickStyles.Clear();
        _tagDefs.Clear();
        _includeTitleHeading = settings.IncludeTitleHeading;
        _offsetHeadings = settings.IncludeTitleHeading;

        var doc = XDocument.Parse(pageXml);
        var page = doc.Root;
        if (page is null)
            return "";

        var title = VisitPage(page);

        var body = _output.ToString().TrimEnd() + "\n";

        var frontmatterProps = new List<string>();

        if (settings.IncludeTitleProperty && !string.IsNullOrWhiteSpace(title))
            frontmatterProps.Add($"title: \"{EscapeYamlString(title!)}\"");

        if (settings.IncludeCreatedDate)
        {
            var created = page.Attribute("dateTime")?.Value;
            if (created is not null && DateTimeOffset.TryParse(created, out var createdDate))
                frontmatterProps.Add($"created: {FormatDate(createdDate, settings.UtcDates)}");
        }

        if (settings.IncludeUpdatedDate)
        {
            var updated = page.Attribute("lastModifiedTime")?.Value;
            if (updated is not null && DateTimeOffset.TryParse(updated, out var updatedDate))
                frontmatterProps.Add($"updated: {FormatDate(updatedDate, settings.UtcDates)}");
        }

        if (frontmatterProps.Count > 0)
        {
            var frontmatter = $"---\n{string.Join('\n', frontmatterProps)}\n---\n";
            return frontmatter + body;
        }

        return body;
    }

    private static string EscapeYamlString(string value)
    {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static string FormatDate(DateTimeOffset date, bool utc)
    {
        var dt = utc ? date.UtcDateTime : date.LocalDateTime;
        var formatted = $"{dt:yyyy-MM-dd'T'HH:mm}:00";
        return utc ? formatted + "Z" : formatted;
    }

    private string? VisitPage(XElement page)
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
        string? titleText = null;
        var title = page.Element(OneNoteNs + "Title");
        if (title is not null)
        {
            titleText = ExtractTextFromOE(title.Element(OneNoteNs + "OE"));
            if (_includeTitleHeading && !string.IsNullOrWhiteSpace(titleText))
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

        return string.IsNullOrWhiteSpace(titleText) ? null : titleText;
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
        var children = oeChildren.Elements().ToList();
        for (var i = 0; i < children.Count; i++)
        {
            var child = children[i];
            if (child.Name.LocalName == "OE")
            {
                if (IsCodeStyle(child))
                {
                    // Collect consecutive code-styled OEs
                    var codeLines = new List<string>();
                    while (i < children.Count && children[i].Name.LocalName == "OE" && IsCodeStyle(children[i]))
                    {
                        codeLines.Add(ExtractTextFromOE(children[i]));
                        i++;
                    }
                    i--; // Back up since for loop will increment

                    if (codeLines.Count == 1 && !codeLines[0].Contains('\n'))
                    {
                        _output.AppendLine($"`{codeLines[0]}`");
                    }
                    else
                    {
                        _output.AppendLine("```");
                        foreach (var line in codeLines)
                            _output.AppendLine(line);
                        _output.AppendLine("```");
                    }
                    _output.AppendLine();
                }
                else
                {
                    VisitOE(child, depth);
                }
            }
            else if (!HandledElements.Contains(child.Name.LocalName))
            {
                _warn($"Unhandled element in OEChildren: <{child.Name.LocalName}>");
            }
        }
    }

    private bool IsCodeStyle(XElement oe)
    {
        var quickStyleIndex = int.Parse(oe.Attribute("quickStyleIndex")?.Value ?? "-1");
        return _quickStyles.TryGetValue(quickStyleIndex, out var style) &&
               style.Name.Equals("code", StringComparison.OrdinalIgnoreCase);
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
        var isTodo = false;
        var isListItem = false;

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
            var effectiveLevel = _offsetHeadings ? headingLevel + 1 : headingLevel;
            var hashes = new string('#', effectiveLevel);
            _output.AppendLine();
            _output.AppendLine($"{hashes} {text}");
            _output.AppendLine();
        }
        else if (isTodo || isListItem)
        {
            var indent = new string(' ', depth * 4);
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
            return WebUtility.HtmlDecode(html);

        var result = new StringBuilder();
        var pos = 0;

        while (pos < html.Length)
        {
            var tagStart = html.IndexOf('<', pos);
            if (tagStart < 0)
            {
                result.Append(WebUtility.HtmlDecode(html[pos..]));
                break;
            }

            // Append text before the tag
            if (tagStart > pos)
                result.Append(WebUtility.HtmlDecode(html[pos..tagStart]));

            var tagEnd = html.IndexOf('>', tagStart);
            if (tagEnd < 0)
            {
                result.Append(html, pos, html.Length - pos);
                break;
            }

            var tag = html.AsSpan(tagStart, tagEnd - tagStart + 1);

            if (StartsWithTag(tag, "span"))
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
            else if (StartsWithTag(tag, "br"))
            {
                result.Append("  \n");
                pos = tagEnd + 1;
                continue;
            }
            else if (StartsWithTag(tag, "a"))
            {
                var href = ExtractAttribute(tag, "href");
                var anchorEnd = html.IndexOf("</a>", tagEnd, StringComparison.OrdinalIgnoreCase);
                if (anchorEnd >= 0)
                {
                    var linkText = ConvertInlineHtml(html[(tagEnd + 1)..anchorEnd]);
                    var hrefStr = WebUtility.HtmlDecode(href.ToString());
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

    private static bool StartsWithTag(ReadOnlySpan<char> tag, string tagName)
    {
        if (!tag.StartsWith($"<{tagName}".AsSpan(), StringComparison.OrdinalIgnoreCase))
            return false;
        var nextCharIndex = 1 + tagName.Length;
        return nextCharIndex >= tag.Length || !char.IsLetter(tag[nextCharIndex]);
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
