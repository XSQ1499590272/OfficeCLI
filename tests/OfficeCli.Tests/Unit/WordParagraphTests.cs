using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordParagraphTests : WordTestBase
{
    [Fact]
    public void AddParagraph_ReadsBackCanonicalParagraphFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Paragraph baseline",
            ["align"] = "center",
            ["spaceBefore"] = "12pt",
            ["spaceAfter"] = "6pt",
            ["lineSpacing"] = "1.5x"
        });

        var node = handler.Get(paraPath);

        Assert.Equal("paragraph", node.Type);
        Assert.Equal("Paragraph baseline", node.Text);
        Assert.Equal("center", Fmt(node)["align"]);
        Assert.Equal("12pt", Fmt(node)["spaceBefore"]);
        Assert.Equal("6pt", Fmt(node)["spaceAfter"]);
        Assert.Equal("1.5x", Fmt(node)["lineSpacing"]);
    }

    [Fact]
    public void SetParagraph_UpdatesCanonicalParagraphFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Mutable" });

        handler.Set(paraPath, new()
        {
            ["align"] = "right",
            ["spaceBefore"] = "3pt",
            ["spaceAfter"] = "9pt",
            ["lineSpacing"] = "2x"
        });

        var node = handler.Get(paraPath);
        Assert.Equal("right", Fmt(node)["align"]);
        Assert.Equal("3pt", Fmt(node)["spaceBefore"]);
        Assert.Equal("9pt", Fmt(node)["spaceAfter"]);
        Assert.Equal("2x", Fmt(node)["lineSpacing"]);
    }
}
