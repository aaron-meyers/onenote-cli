using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
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

    private static readonly Converter ReverseMarkdownConverter = new(new Config
    {
        UnknownTags = Config.UnknownTagsOption.Bypass,
        GithubFlavored = true,
        SmartHrefHandling = false,
    });

    private static string ConvertInlineHtml(string html)
    {
        if (!html.Contains('<'))
            return System.Net.WebUtility.HtmlDecode(html);

        // First convert formatting spans (bold, italic, strikethrough) to semantic tags.
        // This eliminates nested <span> tags that would confuse highlight extraction.
        var normalized = NormalizeSpanStyles(html);

        // Handle highlight spans — now safe because inner spans are already semantic tags.
        // Recursively convert inner content, then wrap with == markers.
        normalized = HighlightSpanRegex().Replace(normalized, match =>
        {
            var inner = ConvertInlineHtml(match.Groups[1].Value);
            return $"=={inner}==";
        });

        // Handle superscript/subscript (vertical-align) spans. Convert them to placeholder
        // tokens so ReverseMarkdown does not strip them, then restore <sup>/<sub> at the end.
        normalized = SuperscriptSpanRegex().Replace(normalized, match =>
        {
            var inner = ConvertInlineHtml(match.Groups[1].Value);
            return $"{SupOpenPlaceholder}{inner}{SupClosePlaceholder}";
        });
        normalized = SubscriptSpanRegex().Replace(normalized, match =>
        {
            var inner = ConvertInlineHtml(match.Groups[1].Value);
            return $"{SubOpenPlaceholder}{inner}{SubClosePlaceholder}";
        });

        string result;
        if (!normalized.Contains('<'))
        {
            result = normalized;
        }
        else
        {
            var markdown = ReverseMarkdownConverter.Convert(normalized);

            // ReverseMarkdown preserves HTML entities in href values; decode them.
            markdown = MarkdownLinkHrefRegex().Replace(markdown, match =>
            {
                var text = match.Groups[1].Value;
                var href = System.Net.WebUtility.HtmlDecode(match.Groups[2].Value);
                return $"[{text}]({href})";
            });

            // ReverseMarkdown may add trailing newlines; trim for inline use
            result = markdown.TrimEnd('\r', '\n');
        }

        return RestoreSupSubPlaceholders(result);
    }

    private const string SupOpenPlaceholder = "\uE000sup\uE000";
    private const string SupClosePlaceholder = "\uE000/sup\uE000";
    private const string SubOpenPlaceholder = "\uE001sub\uE001";
    private const string SubClosePlaceholder = "\uE001/sub\uE001";

    private static string RestoreSupSubPlaceholders(string text)
    {
        return text
            .Replace(SupOpenPlaceholder, "<sup>")
            .Replace(SupClosePlaceholder, "</sup>")
            .Replace(SubOpenPlaceholder, "<sub>")
            .Replace(SubClosePlaceholder, "</sub>");
    }

    [GeneratedRegex(
        @"<span\s[^>]*style\s*=\s*['""][^'""]*vertical-align:\s*super[^'""]*['""][^>]*>([^<]*(?:<(?!/?span[\s>/])[^<]*)*)</span>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex SuperscriptSpanRegex();

    [GeneratedRegex(
        @"<span\s[^>]*style\s*=\s*['""][^'""]*vertical-align:\s*sub[^'""]*['""][^>]*>([^<]*(?:<(?!/?span[\s>/])[^<]*)*)</span>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex SubscriptSpanRegex();

    [GeneratedRegex(
        @"<span\s[^>]*style\s*=\s*""([^""]*)""[^>]*>([^<]*(?:<(?!/?span[\s>/])[^<]*)*)</span>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex SpanStyleRegex();

    [GeneratedRegex(
        @"<span\s[^>]*style\s*=\s*""[^""]*(?:background-color|background:)[^""]*""[^>]*>([^<]*(?:<(?!/?span[\s>/])[^<]*)*)</span>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex HighlightSpanRegex();

    [GeneratedRegex(@"\[([^\]]*)\]\(([^)]*)\)")]
    private static partial Regex MarkdownLinkHrefRegex();

    private static string NormalizeSpanStyles(string html)
    {
        // Repeatedly replace styled spans with semantic HTML tags until none remain
        string previous;
        var result = html;
        do
        {
            previous = result;
            result = SpanStyleRegex().Replace(previous, match =>
            {
                var style = match.Groups[1].Value.ToLowerInvariant();
                var inner = match.Groups[2].Value;

                var open = new StringBuilder();
                var close = new StringBuilder();

                if (style.Contains("font-weight:bold") || style.Contains("font-weight: bold"))
                {
                    open.Append("<strong>");
                    close.Insert(0, "</strong>");
                }

                if (style.Contains("font-style:italic") || style.Contains("font-style: italic"))
                {
                    open.Append("<em>");
                    close.Insert(0, "</em>");
                }

                if (style.Contains("text-decoration:line-through") || style.Contains("text-decoration: line-through"))
                {
                    open.Append("<del>");
                    close.Insert(0, "</del>");
                }

                if (open.Length == 0)
                    return match.Value; // unrecognized or highlight style — keep as-is

                return $"{open}{inner}{close}";
            });
        } while (result != previous);

        return result;
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
