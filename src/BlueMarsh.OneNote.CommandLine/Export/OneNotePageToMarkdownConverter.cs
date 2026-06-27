using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using HtmlAgilityPack;
using ReverseMarkdown;

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

    /// <summary>The line ending used in the output. Defaults to LF ("\n").</summary>
    public string NewLine { get; init; } = "\n";
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
            return NormalizeLineEndings(frontmatter + body, settings.NewLine);
        }

        return NormalizeLineEndings(body, settings.NewLine);
    }

    private static string NormalizeLineEndings(string text, string newLine)
    {
        var normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
        return newLine == "\n" ? normalized : normalized.Replace("\n", newLine);
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
            var name = tagDef.Attribute("name")?.Value ?? "";
            _tagDefs[index] = new TagDefInfo(index, type, symbol, name);
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

        // Check for tag (To-Do or a named label tag)
        var tag = oe.Element(OneNoteNs + "Tag");
        var tagSuffix = "";
        if (tag is not null)
        {
            var tagIndex = int.Parse(tag.Attribute("index")?.Value ?? "-1");
            var completed = tag.Attribute("completed")?.Value == "true";
            if (_tagDefs.TryGetValue(tagIndex, out var tagDef))
            {
                if (IsToDoTag(tagDef))
                {
                    isTodo = true;
                    prefix = completed ? "- [x] " : "- [ ] ";
                }
                else
                {
                    // Non-To-Do tag: append its name as a lowercased hashtag (e.g. #important).
                    tagSuffix = TagHashtagSuffix(tagDef);
                }
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

        // Append any non-To-Do tag hashtag to the line it appears on.
        if (tagSuffix.Length > 0)
            text = (text + tagSuffix).Trim();

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
                    var text = (ExtractTextFromOE(oe) + TagSuffixFor(oe)).Trim();
                    if (!string.IsNullOrEmpty(text))
                        parts.Add(text);
                }
                // Join multi-line cell content with <br>
                return string.Join("<br>", parts);
            }).ToList();

            _output.Append('|');
            foreach (var text in cellTexts)
            {
                _output.Append($" {FormatTableCell(text)} |");
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

    // Collapse any run of line breaks (with surrounding spaces/tabs) within a table cell
    // into a single <br> so the cell stays on one Markdown line, then escape pipes.
    private static string FormatTableCell(string text)
    {
        var collapsed = TableCellNewlineRegex().Replace(text, "<br>").Trim();
        if (collapsed.StartsWith("<br>", StringComparison.Ordinal))
            collapsed = collapsed[4..];
        if (collapsed.EndsWith("<br>", StringComparison.Ordinal))
            collapsed = collapsed[..^4];
        return collapsed.Replace("|", "\\|");
    }

    [GeneratedRegex(@"[ \t]*(?:\r\n|\r|\n)+[ \t]*")]
    private static partial Regex TableCellNewlineRegex();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();

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

    private static readonly Converter ReverseMarkdownConverter = CreateReverseMarkdownConverter();

    private static Converter CreateReverseMarkdownConverter()
    {
        var config = new Config
        {
            UnknownTags = Config.UnknownTagsOption.Bypass,
            GithubFlavored = true,
            SmartHrefHandling = false,
        };

        // Superscript/subscript have no Markdown equivalent, so allow them through as raw HTML.
        config.PassThroughTags.Add("sup");
        config.PassThroughTags.Add("sub");

        // Highlighted text is rendered with == markers (e.g. ==text==). The replacer wraps the
        // converted inner content with the given string as both prefix and suffix.
        config.UnknownTagsReplacer["mark"] = "==";

        return new Converter(config);
    }

    private static string ConvertInlineHtml(string html)
    {
        if (!html.Contains('<'))
            return System.Net.WebUtility.HtmlDecode(html);

        // Rewrite OneNote's inline-styled <span> elements into semantic tags
        // (<mark>, <strong>, <em>, <del>, <sup>, <sub>) that ReverseMarkdown understands.
        var normalized = NormalizeHtml(html);

        if (!normalized.Contains('<'))
        {
            // No tags remain (e.g. a color-only span was unwrapped); still decode any HTML
            // entities that were part of the surrounding text.
            return System.Net.WebUtility.HtmlDecode(normalized);
        }

        var markdown = ReverseMarkdownConverter.Convert(normalized);

        // ReverseMarkdown may add trailing newlines; trim for inline use. Decode any HTML
        // entities it preserved (e.g. &lt; &gt; around autolinks, and entities inside hrefs).
        return System.Net.WebUtility.HtmlDecode(markdown.TrimEnd('\r', '\n'));
    }

    // Parses the HTML fragment and rewrites OneNote's inline-styled <span> elements into the
    // equivalent semantic HTML tags. Spans whose styles aren't recognized (e.g. a bare color)
    // are unwrapped, leaving their content in place.
    private static string NormalizeHtml(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        NormalizeSpans(doc.DocumentNode);
        return doc.DocumentNode.InnerHtml;
    }

    private static void NormalizeSpans(HtmlNode node)
    {
        // Process children first; a styled span may be nested inside another span.
        foreach (var child in node.ChildNodes.ToList())
            NormalizeSpans(child);

        if (node.NodeType != HtmlNodeType.Element ||
            !node.Name.Equals("span", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var tags = SemanticTagsForStyle(node.GetAttributeValue("style", ""));
        var parent = node.ParentNode;
        var children = node.ChildNodes.ToList();

        if (tags.Count == 0)
        {
            // Unrecognized style (e.g. color only) — unwrap, keeping the span's content.
            foreach (var child in children)
                parent.InsertBefore(child, node);
            parent.RemoveChild(node);
            return;
        }

        // Wrap the span's content in nested semantic elements (first tag is outermost).
        var doc = node.OwnerDocument;
        var outer = doc.CreateElement(tags[0]);
        var innermost = outer;
        for (var i = 1; i < tags.Count; i++)
        {
            var inner = doc.CreateElement(tags[i]);
            innermost.AppendChild(inner);
            innermost = inner;
        }

        foreach (var child in children)
            innermost.AppendChild(child);

        parent.ReplaceChild(outer, node);
    }

    // Maps a CSS style string to the ordered list of semantic HTML tags (outermost first)
    // that represent it. Returns an empty list if no formatting is recognized.
    private static List<string> SemanticTagsForStyle(string style)
    {
        var s = WhitespaceRegex().Replace(style, "").ToLowerInvariant();
        var tags = new List<string>();

        if (s.Contains("background-color:") || s.Contains("background:"))
            tags.Add("mark");
        if (s.Contains("font-weight:bold"))
            tags.Add("strong");
        if (s.Contains("font-style:italic"))
            tags.Add("em");
        if (s.Contains("text-decoration:line-through"))
            tags.Add("del");
        if (s.Contains("vertical-align:super"))
            tags.Add("sup");
        if (s.Contains("vertical-align:sub"))
            tags.Add("sub");

        return tags;
    }

    private static bool IsToDoTag(TagDefInfo tagDef)
    {
        // To-Do tag: type="0" symbol="3" or type="0" symbol="0" (common patterns)
        // Also check for type="0" which is the standard checkbox
        return tagDef.Type == "0";
    }

    // Converts a tag name to a lowercased hashtag (e.g. "Important" -> "#important",
    // "To Do" -> "#todo"). Returns null if no usable characters remain.
    private static string? ToHashtag(string tagName)
    {
        var sb = new StringBuilder();
        foreach (var c in tagName)
        {
            if (char.IsLetterOrDigit(c))
                sb.Append(char.ToLowerInvariant(c));
        }
        return sb.Length == 0 ? null : "#" + sb;
    }

    // Returns " #hashtag" for a non-To-Do tag definition, or "" if not applicable.
    private static string TagHashtagSuffix(TagDefInfo tagDef)
    {
        if (IsToDoTag(tagDef))
            return "";
        var hashtag = ToHashtag(tagDef.Name);
        return hashtag is null ? "" : $" {hashtag}";
    }

    // Resolves the non-To-Do tag (if any) on an OE element to its " #hashtag" suffix.
    private string TagSuffixFor(XElement oe)
    {
        var tag = oe.Element(OneNoteNs + "Tag");
        if (tag is null)
            return "";
        var tagIndex = int.Parse(tag.Attribute("index")?.Value ?? "-1");
        return _tagDefs.TryGetValue(tagIndex, out var tagDef) ? TagHashtagSuffix(tagDef) : "";
    }

    private sealed record QuickStyleInfo(int Index, string Name);
    private sealed record TagDefInfo(int Index, string Type, string Symbol, string Name);
}
