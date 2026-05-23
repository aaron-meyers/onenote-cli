using BlueMarsh.OneNote.CommandLine.Export;

namespace BlueMarsh.OneNote.CommandLine.Tests;

public class OneNotePageToMarkdownConverterTests
{
    private static readonly string OneNoteNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private static string WrapPage(string body, string? styles = null) =>
        $"""
        <?xml version="1.0" encoding="utf-8"?>
        <one:Page xmlns:one="{OneNoteNs}">
          {styles}
          <one:Title>
            <one:OE>
              <one:T><![CDATA[Test Page]]></one:T>
            </one:OE>
          </one:Title>
          <one:Outline>
            <one:OEChildren>
              {body}
            </one:OEChildren>
          </one:Outline>
        </one:Page>
        """;

    private static string Convert(string pageXml)
    {
        var warnings = new List<string>();
        var converter = new OneNotePageToMarkdownConverter(w => warnings.Add(w));
        return converter.Convert(pageXml, new MarkdownConversionSettings { IncludeTitleHeading = true });
    }

    [Test]
    public Task PlainText()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Hello, world!]]></one:T>
            </one:OE>
            <one:OE>
              <one:T><![CDATA[Second paragraph.]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task Headings()
    {
        var xml = WrapPage("""
            <one:OE quickStyleIndex="1">
              <one:T><![CDATA[Heading 1]]></one:T>
            </one:OE>
            <one:OE quickStyleIndex="2">
              <one:T><![CDATA[Heading 2]]></one:T>
            </one:OE>
            <one:OE quickStyleIndex="3">
              <one:T><![CDATA[Heading 3]]></one:T>
            </one:OE>
            """,
            styles: """
            <one:QuickStyleDef index="1" name="h1" />
            <one:QuickStyleDef index="2" name="h2" />
            <one:QuickStyleDef index="3" name="h3" />
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task BulletList()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:List><one:Bullet bullet="2" /></one:List>
              <one:T><![CDATA[First item]]></one:T>
            </one:OE>
            <one:OE>
              <one:List><one:Bullet bullet="2" /></one:List>
              <one:T><![CDATA[Second item]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task NumberedList()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:List><one:Number numberSequence="0" numberFormat="##." /></one:List>
              <one:T><![CDATA[Step one]]></one:T>
            </one:OE>
            <one:OE>
              <one:List><one:Number numberSequence="0" numberFormat="##." /></one:List>
              <one:T><![CDATA[Step two]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task TodoItems()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:Tag index="0" completed="false" />
              <one:T><![CDATA[Incomplete task]]></one:T>
            </one:OE>
            <one:OE>
              <one:Tag index="0" completed="true" />
              <one:T><![CDATA[Completed task]]></one:T>
            </one:OE>
            """,
            styles: """
            <one:TagDef index="0" type="0" symbol="0" />
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task BoldItalicStrikethrough()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Normal <span style="font-weight:bold">bold</span> normal]]></one:T>
            </one:OE>
            <one:OE>
              <one:T><![CDATA[Normal <span style="font-style:italic">italic</span> normal]]></one:T>
            </one:OE>
            <one:OE>
              <one:T><![CDATA[Normal <span style="text-decoration:line-through">struck</span> normal]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task Links()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Visit <a href="https://example.com">Example</a> for more.]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task LinkWithNewlineInTag()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[See [<a
            href="https://example.com/docs">info</a>] for details.]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task InlineCodeSingleLine()
    {
        var xml = WrapPage("""
            <one:OE quickStyleIndex="1">
              <one:T><![CDATA[echo "hello"]]></one:T>
            </one:OE>
            """,
            styles: """<one:QuickStyleDef index="1" name="code" />""");

        return Verify(Convert(xml));
    }

    [Test]
    public Task FencedCodeBlock()
    {
        var xml = WrapPage("""
            <one:OE quickStyleIndex="1">
              <one:T><![CDATA[def greet():]]></one:T>
            </one:OE>
            <one:OE quickStyleIndex="1">
              <one:T><![CDATA[    print("hello")]]></one:T>
            </one:OE>
            <one:OE quickStyleIndex="1">
              <one:T><![CDATA[    print("world")]]></one:T>
            </one:OE>
            """,
            styles: """<one:QuickStyleDef index="1" name="code" />""");

        return Verify(Convert(xml));
    }

    [Test]
    public Task Table()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:Table>
                <one:Row>
                  <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Name]]></one:T></one:OE></one:OEChildren></one:Cell>
                  <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Value]]></one:T></one:OE></one:OEChildren></one:Cell>
                </one:Row>
                <one:Row>
                  <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Alpha]]></one:T></one:OE></one:OEChildren></one:Cell>
                  <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[100]]></one:T></one:OE></one:OEChildren></one:Cell>
                </one:Row>
              </one:Table>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task NestedList()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:List><one:Bullet bullet="2" /></one:List>
              <one:T><![CDATA[Parent]]></one:T>
              <one:OEChildren>
                <one:OE>
                  <one:List><one:Bullet bullet="2" /></one:List>
                  <one:T><![CDATA[Child 1]]></one:T>
                </one:OE>
                <one:OE>
                  <one:List><one:Bullet bullet="2" /></one:List>
                  <one:T><![CDATA[Child 2]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task NoTitleHeading()
    {
        var xml = WrapPage("""
            <one:OE quickStyleIndex="1">
              <one:T><![CDATA[Heading 1]]></one:T>
            </one:OE>
            <one:OE>
              <one:T><![CDATA[Body text.]]></one:T>
            </one:OE>
            """,
            styles: """<one:QuickStyleDef index="1" name="h1" />""");

        var converter = new OneNotePageToMarkdownConverter(w => { });
        return Verify(converter.Convert(xml));
    }

    [Test]
    public Task TitleProperty()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Body text.]]></one:T>
            </one:OE>
            """);

        var converter = new OneNotePageToMarkdownConverter(w => { });
        return Verify(converter.Convert(xml, new MarkdownConversionSettings { IncludeTitleProperty = true }));
    }

    [Test]
    public Task TitlePropertyWithSpecialCharacters()
    {
        var pageXml = $"""
            <?xml version="1.0" encoding="utf-8"?>
            <one:Page xmlns:one="{OneNoteNs}">
              <one:Title>
                <one:OE>
                  <one:T><![CDATA[Notes: "Project A/B"]]></one:T>
                </one:OE>
              </one:Title>
              <one:Outline>
                <one:OEChildren>
                  <one:OE>
                    <one:T><![CDATA[Content here.]]></one:T>
                  </one:OE>
                </one:OEChildren>
              </one:Outline>
            </one:Page>
            """;

        var converter = new OneNotePageToMarkdownConverter(w => { });
        return Verify(converter.Convert(pageXml, new MarkdownConversionSettings { IncludeTitleProperty = true }));
    }

    [Test]
    public Task TitlePropertyAndHeading()
    {
        var xml = WrapPage("""
            <one:OE quickStyleIndex="1">
              <one:T><![CDATA[Section]]></one:T>
            </one:OE>
            <one:OE>
              <one:T><![CDATA[Body text.]]></one:T>
            </one:OE>
            """,
            styles: """<one:QuickStyleDef index="1" name="h1" />""");

        var converter = new OneNotePageToMarkdownConverter(w => { });
        return Verify(converter.Convert(xml, new MarkdownConversionSettings { IncludeTitleProperty = true, IncludeTitleHeading = true }));
    }

    [Test]
    public Task CreatedDate()
    {
        var pageXml = $"""
            <?xml version="1.0" encoding="utf-8"?>
            <one:Page xmlns:one="{OneNoteNs}" dateTime="2025-03-15T09:30:00.000Z" lastModifiedTime="2025-04-20T14:45:00.000Z">
              <one:Title>
                <one:OE>
                  <one:T><![CDATA[Test Page]]></one:T>
                </one:OE>
              </one:Title>
              <one:Outline>
                <one:OEChildren>
                  <one:OE>
                    <one:T><![CDATA[Body text.]]></one:T>
                  </one:OE>
                </one:OEChildren>
              </one:Outline>
            </one:Page>
            """;

        var converter = new OneNotePageToMarkdownConverter(w => { });
        return Verify(converter.Convert(pageXml, new MarkdownConversionSettings { IncludeCreatedDate = true, UtcDates = true }));
    }

    [Test]
    public Task UpdatedDate()
    {
        var pageXml = $"""
            <?xml version="1.0" encoding="utf-8"?>
            <one:Page xmlns:one="{OneNoteNs}" dateTime="2025-03-15T09:30:00.000Z" lastModifiedTime="2025-04-20T14:45:00.000Z">
              <one:Title>
                <one:OE>
                  <one:T><![CDATA[Test Page]]></one:T>
                </one:OE>
              </one:Title>
              <one:Outline>
                <one:OEChildren>
                  <one:OE>
                    <one:T><![CDATA[Body text.]]></one:T>
                  </one:OE>
                </one:OEChildren>
              </one:Outline>
            </one:Page>
            """;

        var converter = new OneNotePageToMarkdownConverter(w => { });
        return Verify(converter.Convert(pageXml, new MarkdownConversionSettings { IncludeUpdatedDate = true, UtcDates = true }));
    }

    [Test]
    public Task HtmlEntityDecoding()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Strings &amp; things &quot;quoted&quot; with &lt;angle brackets&gt; and &#39;apostrophes&#39;]]></one:T>
            </one:OE>
            <one:OE>
              <one:T><![CDATA[<a href="https://example.com?foo=1&amp;bar=2">link &amp; text</a>]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task BothDatesWithTitle()
    {
        var pageXml = $"""
            <?xml version="1.0" encoding="utf-8"?>
            <one:Page xmlns:one="{OneNoteNs}" dateTime="2025-03-15T09:30:00.000Z" lastModifiedTime="2025-04-20T14:45:00.000Z">
              <one:Title>
                <one:OE>
                  <one:T><![CDATA[Test Page]]></one:T>
                </one:OE>
              </one:Title>
              <one:Outline>
                <one:OEChildren>
                  <one:OE>
                    <one:T><![CDATA[Body text.]]></one:T>
                  </one:OE>
                </one:OEChildren>
              </one:Outline>
            </one:Page>
            """;

        var converter = new OneNotePageToMarkdownConverter(w => { });
        return Verify(converter.Convert(pageXml, new MarkdownConversionSettings
        {
            IncludeTitleProperty = true,
            IncludeCreatedDate = true,
            IncludeUpdatedDate = true,
            UtcDates = true,
        }));
    }

    [Test]
    public Task BoldItalicTags()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Normal <b>bold</b> and <i>italic</i> normal]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task StrikethroughTag()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Normal <s>struck</s> and <strike>also struck</strike> normal]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task Highlight()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Normal <span style="background-color:yellow">highlighted</span> normal]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task HighlightWithFormatting()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[<span style="background-color:yellow"><span style="font-weight:bold">bold highlight</span></span>]]></one:T>
            </one:OE>
            <one:OE>
              <one:T><![CDATA[<span style="background-color:cyan"><span style="font-style:italic">italic highlight</span></span>]]></one:T>
            </one:OE>
            <one:OE>
              <one:T><![CDATA[<span style="background-color:red"><span style="text-decoration:line-through">struck highlight</span></span>]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task NestedFormatting()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[<b><i>bold italic</i></b> and <span style="font-weight:bold"><span style="font-style:italic">also bold italic</span></span>]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }

    [Test]
    public Task ColorSpanStripped()
    {
        var xml = WrapPage("""
            <one:OE>
              <one:T><![CDATA[Normal <span style="color:#ff0000">red text</span> normal]]></one:T>
            </one:OE>
            """);

        return Verify(Convert(xml));
    }
}
