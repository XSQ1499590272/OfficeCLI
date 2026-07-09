using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordRawAndFindReplaceTests : WordTestBase
{
    [Fact]
    public void SetFindReplace_ReplacesPlainTextAndRegexBackreferences()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.Add("/body", "paragraph", null, new() { ["text"] = "needle here" });
        var second = handler.Add("/body", "paragraph", null, new() { ["text"] = "Invoice 42" });

        handler.Set("/body", new()
        {
            ["find"] = "needle",
            ["replace"] = "thread"
        });
        handler.Set(second, new()
        {
            ["find"] = "(\\d+)",
            ["regex"] = "true",
            ["replace"] = "#$1"
        });

        Assert.Equal("thread here", handler.Get(first).Text);
        Assert.Equal("Invoice #42", handler.Get(second).Text);
        Assert.Equal(1, handler.LastFindMatchCount);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedSet_BareFindThrowsWithoutChangingText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "needle" });

        var ex = Assert.Throws<ArgumentException>(() =>
            handler.Set("/body", new() { ["find"] = "needle" }));

        Assert.Contains("'find' requires", ex.Message);
        Assert.Equal("needle", handler.Get(paragraphPath).Text);
    }

    [Fact]
    public void SetFindReplace_CrossRunSucceedsButCrossHyperlinkBoundaryIsRejected()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var splitParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(splitParagraph, "run", null, new() { ["text"] = "foo" });
        handler.Add(splitParagraph, "run", null, new() { ["text"] = "bar" });
        var linkedParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var linkPath = handler.Add(linkedParagraph, "hyperlink", null, new()
        {
            ["url"] = "https://example.com",
            ["text"] = "link"
        });
        handler.Add(linkedParagraph, "run", null, new() { ["text"] = "tail" });

        handler.Set(splitParagraph, new()
        {
            ["find"] = "foobar",
            ["replace"] = "joined"
        });
        var ex = Assert.Throws<ArgumentException>(() =>
            handler.Set(linkedParagraph, new()
            {
                ["find"] = "linktail",
                ["replace"] = "broken"
            }));

        Assert.Equal("joined", handler.Get(splitParagraph).Text);
        Assert.Contains("cannot span a hyperlink boundary", ex.Message);
        Assert.Equal("linktail", handler.Get(linkedParagraph).Text);
        Assert.Equal("https://example.com", Fmt(handler.Get(linkPath))["url"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void RawAndRawSet_ReadWriteAndFailedXPathDoesNotMutateDocument()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Raw target" });
        var before = handler.Raw("/document");

        Assert.Contains("<w:document", before);
        Assert.Contains("<w:styles", handler.Raw("/styles"));

        handler.RawSet("/document", "//w:p[1]", "setattr", "w:rsidR=00112233");
        Assert.Contains("00112233", handler.Raw("/document"));

        var afterSuccess = handler.Raw("/document");
        var ex = Assert.Throws<ArgumentException>(() =>
            handler.RawSet("/document", "//w:p[99]", "setattr", "w:rsidR=44556677"));

        Assert.Contains("XPath matched no elements", ex.Message);
        Assert.Equal(afterSuccess, handler.Raw("/document"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedRawSet_InvalidXmlFragmentDoesNotMutateDocument()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Raw target" });
        var before = handler.Raw("/document");

        var ex = Assert.Throws<System.Xml.XmlException>(() =>
            handler.RawSet("/document", "//w:body", "append", "<w:p><w:r>"));

        Assert.Contains("w:r", ex.Message);
        Assert.Equal(before, handler.Raw("/document"));
        Assert.Empty(handler.Validate());
    }
}
