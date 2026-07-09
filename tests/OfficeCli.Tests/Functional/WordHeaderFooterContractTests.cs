using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordHeaderFooterContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddHeader_ReadsBackTextTypeFormattingAndChildParagraph()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var headerPath = handler.Add("/", "header", null, new()
        {
            ["type"] = "default",
            ["text"] = "Header text",
            ["align"] = "center",
            ["font"] = "Arial",
            ["size"] = "12",
            ["bold"] = "true",
            ["color"] = "2A9D8F"
        });

        var header = handler.Get(headerPath, depth: 1);
        var queried = Assert.Single(handler.Query("header:contains(\"Header\")"));

        Assert.Equal("header", header.Type);
        Assert.Equal("Header text", header.Text);
        Assert.Equal("default", Fmt(header)["type"]);
        Assert.Equal("center", Fmt(header)["align"]);
        Assert.Equal("Arial", Fmt(header)["font"]);
        Assert.Equal("12pt", Fmt(header)["size"]);
        Assert.Equal(true, Fmt(header)["bold"]);
        Assert.Equal("#2A9D8F", Fmt(header)["color"]);
        Assert.Single(header.Children);
        Assert.Equal("paragraph", header.Children[0].Type);
        Assert.Equal(headerPath, queried.Path);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddFooter_ReadsBackTextTypeFormattingAndAllowsChildParagraphAdd()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var footerPath = handler.Add("/", "footer", null, new()
        {
            ["type"] = "first",
            ["text"] = "Footer text",
            ["align"] = "right",
            ["italic"] = "true"
        });
        var paragraphPath = handler.Add(footerPath, "paragraph", null, new() { ["text"] = "Second footer line" });

        var footer = handler.Get(footerPath, depth: 1);
        var paragraph = handler.Get(paragraphPath);
        var queried = Assert.Single(handler.Query("footer:contains(\"Footer\")"));

        Assert.Equal("footer", footer.Type);
        Assert.Contains("Footer text", footer.Text);
        Assert.Equal("first", Fmt(footer)["type"]);
        Assert.Equal("right", Fmt(footer)["align"]);
        Assert.Equal(true, Fmt(footer)["italic"]);
        Assert.Equal("paragraph", paragraph.Type);
        Assert.Equal("Second footer line", paragraph.Text);
        Assert.Equal(footerPath, queried.Path);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedAdd_DuplicateDefaultHeader_DoesNotCreateSecondHeader()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/", "header", null, new() { ["text"] = "First header" });

            Assert.Throws<ArgumentException>(() =>
                handler.Add("/", "header", null, new() { ["text"] = "Duplicate header" }));
        }

        using var reopened = new WordHandler(path, editable: false);
        Assert.Single(reopened.Query("header"));
        Assert.Equal("First header", reopened.Query("header")[0].Text);
    }
}
