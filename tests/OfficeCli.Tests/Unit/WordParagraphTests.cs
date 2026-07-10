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

    [Fact]
    public void SetParagraph_ListStyleAndIndentReadBackCurrentConflictBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Indented item" });

        var unsupported = handler.Set(paraPath, new()
        {
            ["firstLineIndent"] = "12pt",
            ["rightIndent"] = "18pt",
            ["hangingIndent"] = "6pt",
            ["listStyle"] = "bullet"
        });

        var node = handler.Get(paraPath);
        Assert.Empty(unsupported);
        Assert.False(Fmt(node).ContainsKey("firstLineIndent"));
        Assert.Equal("18pt", Fmt(node)["rightIndent"]);
        Assert.Equal("6pt", Fmt(node)["hangingIndent"]);
        Assert.Equal("bullet", Fmt(node)["listStyle"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddParagraph_ReadsBackPaginationDirectionAndPatternShading()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "advanced paragraph",
            ["keepNext"] = "true",
            ["keepLines"] = "true",
            ["pageBreakBefore"] = "true",
            ["widowControl"] = "false",
            ["bidi"] = "rtl",
            ["shading"] = "pct20;FFFF00;0000FF"
        });

        var format = Fmt(handler.Get(paragraphPath));

        Assert.Equal(true, format["keepNext"]);
        Assert.Equal(true, format["keepLines"]);
        Assert.Equal(true, format["pageBreakBefore"]);
        Assert.Equal(false, format["widowControl"]);
        Assert.Equal("rtl", format["direction"]);
        Assert.Equal("pct20", format["shading.val"]);
        Assert.Equal("#FFFF00", format["shading.fill"]);
        Assert.Equal("#0000FF", format["shading.color"]);
        Assert.Empty(handler.Validate());
    }
}
