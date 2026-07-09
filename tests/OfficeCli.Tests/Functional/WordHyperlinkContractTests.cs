using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordHyperlinkContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddHyperlink_ReadsBackUrlTextAndFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var hyperlinkPath = handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["url"] = "https://example.com/docs",
            ["text"] = "Example docs",
            ["tooltip"] = "Open docs",
            ["color"] = "C00000",
            ["underline"] = "single"
        });

        var hyperlink = handler.Get(hyperlinkPath);

        Assert.Equal("hyperlink", hyperlink.Type);
        Assert.Equal("Example docs", hyperlink.Text);
        Assert.Equal("https://example.com/docs", Fmt(hyperlink)["url"]);
        Assert.Equal("Open docs", Fmt(hyperlink)["tooltip"]);
        Assert.Equal("#C00000", Fmt(hyperlink)["color"]);
        Assert.Equal("single", Fmt(hyperlink)["underline"]);
    }

    [Fact]
    public void AddHyperlink_FragmentUrlPromotesToAnchor()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var hyperlinkPath = handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["url"] = "#TargetBookmark",
            ["text"] = "Jump"
        });

        var hyperlink = handler.Get(hyperlinkPath);

        Assert.Equal("hyperlink", hyperlink.Type);
        Assert.Equal("Jump", hyperlink.Text);
        Assert.Equal("TargetBookmark", Fmt(hyperlink)["anchor"]);
        Assert.False(Fmt(hyperlink).ContainsKey("url"));
    }
}
