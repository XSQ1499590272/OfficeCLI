using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordTextInputBoundaryTests : WordTestBase
{
    [Fact]
    public void AddParagraph_TextWithNewlineTabAndPageTokens_UsesStructuralOpenXml()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            var paragraphPath = handler.Add("/body", "paragraph", null, new()
            {
                ["text"] = "A\nB\tC page {page} of {pages}"
            });

            Assert.Equal("A\nB\tC page 1 of 1", handler.Get(paragraphPath).Text);
        }

        using var doc = WordprocessingDocument.Open(path, false);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single();
        Assert.Single(paragraph.Descendants<Break>());
        Assert.Single(paragraph.Descendants<TabChar>());
        Assert.Equal(new[] { " PAGE ", " NUMPAGES " },
            paragraph.Descendants<FieldCode>().Select(f => f.Text!).ToArray());
    }

    [Fact]
    public void AddParagraph_IllegalXmlControlCharacter_IsRejectedWithoutAddingParagraph()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        Assert.Throws<ArgumentException>(() =>
            handler.Add("/body", "paragraph", null, new() { ["text"] = "bad\u0001text" }));

        Assert.Empty(handler.Query("paragraph"));
    }

    [Fact]
    public void TextEscape_ResolvesSupportedEscapesInOnePass()
    {
        Assert.Equal("line1\nline2\tvalue\rreturn\\", TextEscape.Resolve("line1\\nline2\\tvalue\\rreturn\\\\"));
        Assert.Equal("literal\\n", TextEscape.Resolve("literal\\\\n"));
    }

    [Fact]
    public void TextEscape_PreservesUnknownAndTrailingBackslashSequences()
    {
        Assert.Equal("unknown\\q", TextEscape.Resolve("unknown\\q"));
        Assert.Equal("trailing\\", TextEscape.Resolve("trailing\\"));
        Assert.Equal("plain", TextEscape.Resolve("plain"));
    }

    [Fact]
    public void TextEscape_TreatsNullAndEmptyAsEmpty()
    {
        Assert.Equal(string.Empty, TextEscape.Resolve(null));
        Assert.Equal(string.Empty, TextEscape.Resolve(string.Empty));
    }
}
