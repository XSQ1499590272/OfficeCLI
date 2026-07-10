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

    [Fact]
    public void AddHyperlink_ReadsBackWrapperMetadataAndRunFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var hyperlinkPath = handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["url"] = "https://example.com/report",
            ["text"] = "Styled report",
            ["tooltip"] = "Open report",
            ["tgtFrame"] = "_blank",
            ["history"] = "false",
            ["docLocation"] = "Section 2",
            ["font"] = "Arial",
            ["font.cs"] = "Arial Unicode MS",
            ["size"] = "13",
            ["bold"] = "true",
            ["italic"] = "true",
            ["underline"] = "wave",
            ["underline.color"] = "00AA00",
            ["strike"] = "true",
            ["highlight"] = "yellow",
            ["rStyle"] = "Hyperlink"
        });

        var hyperlink = handler.Get(hyperlinkPath);
        var fmt = Fmt(hyperlink);

        Assert.Equal("Styled report", hyperlink.Text);
        Assert.Equal("https://example.com/report", fmt["url"]);
        Assert.Equal("Open report", fmt["tooltip"]);
        Assert.Equal("_blank", fmt["tgtFrame"]);
        Assert.Equal(false, fmt["history"]);
        Assert.Equal("Section 2", fmt["docLocation"]);
        Assert.Equal("Arial", fmt["font"]);
        Assert.Equal("Arial Unicode MS", fmt["font.cs"]);
        Assert.Equal("13pt", fmt["size"]);
        Assert.Equal(true, fmt["bold"]);
        Assert.Equal(true, fmt["italic"]);
        Assert.Equal("wave", fmt["underline"]);
        Assert.Equal("#00AA00", fmt["underline.color"]);
        Assert.Equal(true, fmt["strike"]);
        Assert.Equal("yellow", fmt["highlight"]);
        Assert.Equal("Hyperlink", fmt["rStyle"]);
    }

    [Fact]
    public void AddHyperlink_ReadsBackThemeColorAndInheritSentinels()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var themedPath = handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["url"] = "https://example.com/themed",
            ["text"] = "Themed",
            ["color"] = "accent1",
            ["underline"] = "dotted"
        });
        var inheritedPath = handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["url"] = "https://example.com/plain",
            ["text"] = "Plain",
            ["color"] = "inherit",
            ["underline"] = "inherit"
        });

        var themed = Fmt(handler.Get(themedPath));
        var inherited = Fmt(handler.Get(inheritedPath));

        Assert.Equal("accent1", themed["color"]);
        Assert.Equal("dotted", themed["underline"]);
        Assert.False(inherited.ContainsKey("color"));
        Assert.False(inherited.ContainsKey("underline"));
    }
}
