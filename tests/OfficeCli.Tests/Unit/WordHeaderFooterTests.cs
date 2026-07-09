using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordHeaderFooterTests : WordTestBase
{
    [Fact]
    public void AddHeaderAndFooter_ReadBackFormattingAndChildParagraphs()
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
        var footerPath = handler.Add("/", "footer", null, new()
        {
            ["type"] = "first",
            ["text"] = "Footer text",
            ["align"] = "right",
            ["italic"] = "true"
        });
        var footerParagraphPath = handler.Add(footerPath, "paragraph", null, new() { ["text"] = "Second footer line" });

        var header = handler.Get(headerPath, depth: 1);
        var footer = handler.Get(footerPath, depth: 1);
        var footerParagraph = handler.Get(footerParagraphPath);

        Assert.Equal("Header text", header.Text);
        Assert.Equal("default", Fmt(header)["type"]);
        Assert.Equal("center", Fmt(header)["align"]);
        Assert.Equal("Arial", Fmt(header)["font"]);
        Assert.Equal("12pt", Fmt(header)["size"]);
        Assert.Equal(true, Fmt(header)["bold"]);
        Assert.Equal("#2A9D8F", Fmt(header)["color"]);
        Assert.Single(header.Children);
        Assert.Contains("Footer text", footer.Text);
        Assert.Contains("Second footer line", footer.Text);
        Assert.Equal("first", Fmt(footer)["type"]);
        Assert.Equal("right", Fmt(footer)["align"]);
        Assert.Equal(true, Fmt(footer)["italic"]);
        Assert.Equal("Second footer line", footerParagraph.Text);
        Assert.Equal(headerPath, Assert.Single(handler.Query("header:contains(\"Header\")")).Path);
        Assert.Equal(footerPath, Assert.Single(handler.Query("footer:contains(\"Footer\")")).Path);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedAdd_DuplicateDefaultHeader_DoesNotCreateSecondHeader()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/", "header", null, new() { ["text"] = "First header" });
            handler.Add("/", "header", null, new() { ["type"] = "first", ["text"] = "First-page header" });
            handler.Add("/", "header", null, new() { ["type"] = "even", ["text"] = "Even-page header" });
            handler.Add("/", "footer", null, new() { ["text"] = "Default footer" });
            handler.Add("/", "footer", null, new() { ["type"] = "first", ["text"] = "First-page footer" });
            handler.Add("/", "footer", null, new() { ["type"] = "even", ["text"] = "Even-page footer" });

            Assert.Throws<ArgumentException>(() =>
                handler.Add("/", "header", null, new() { ["text"] = "Duplicate header" }));
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/", "header", null, new() { ["type"] = "first", ["text"] = "Duplicate first header" }));
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/", "header", null, new() { ["type"] = "even", ["text"] = "Duplicate even header" }));
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/", "footer", null, new() { ["text"] = "Duplicate footer" }));
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/", "footer", null, new() { ["type"] = "first", ["text"] = "Duplicate first footer" }));
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/", "footer", null, new() { ["type"] = "even", ["text"] = "Duplicate even footer" }));
        }

        using var reopened = new WordHandler(path, editable: false);
        var headers = reopened.Query("header");
        var footers = reopened.Query("footer");

        Assert.Equal(3, headers.Count);
        Assert.Equal(3, footers.Count);
        Assert.Contains(headers, h => h.Text == "First header");
        Assert.Contains(headers, h => h.Text == "First-page header");
        Assert.Contains(headers, h => h.Text == "Even-page header");
        Assert.Contains(footers, f => f.Text == "Default footer");
        Assert.Contains(footers, f => f.Text == "First-page footer");
        Assert.Contains(footers, f => f.Text == "Even-page footer");
    }
}
