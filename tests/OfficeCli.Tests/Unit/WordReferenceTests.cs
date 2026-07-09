using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordReferenceTests : WordTestBase
{
    [Fact]
    public void AddHyperlink_ReadsBackExternalAndFragmentAnchorForms()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var externalPath = handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["url"] = "https://example.com/docs",
            ["text"] = "Example docs",
            ["tooltip"] = "Open docs",
            ["color"] = "C00000",
            ["underline"] = "single"
        });
        var anchorPath = handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["url"] = "#TargetBookmark",
            ["text"] = "Jump"
        });

        var external = handler.Get(externalPath);
        var anchor = handler.Get(anchorPath);

        Assert.Equal("Example docs", external.Text);
        Assert.Equal("https://example.com/docs", Fmt(external)["url"]);
        Assert.Equal("Open docs", Fmt(external)["tooltip"]);
        Assert.Equal("#C00000", Fmt(external)["color"]);
        Assert.Equal("single", Fmt(external)["underline"]);
        Assert.Equal("Jump", anchor.Text);
        Assert.Equal("TargetBookmark", Fmt(anchor)["anchor"]);
        Assert.False(Fmt(anchor).ContainsKey("url"));
    }

    [Fact]
    public void AddBookmark_ReadsBackDuplicatesAndRejectsMissingName()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "bookmark host" });
        var firstPath = handler.Add(paragraphPath, "bookmark", null, new() { ["name"] = "DuplicateMark" });
        handler.Add(paragraphPath, "bookmark", null, new() { ["name"] = "DuplicateMark" });
        handler.Add(paragraphPath, "bookmark", null, new() { ["name"] = "Review/Analysis" });

        var first = handler.Get(firstPath);
        var bookmarks = handler.Query("bookmark");

        Assert.Equal("bookmark", first.Type);
        Assert.Equal("DuplicateMark", Fmt(first)["name"]);
        Assert.Equal(2, bookmarks.Count(node => Equals(Fmt(node)["name"], "DuplicateMark")));
        Assert.Contains(bookmarks, node => Equals(Fmt(node)["name"], "Review/Analysis"));
        Assert.Throws<ArgumentException>(() => handler.Add(paragraphPath, "bookmark", null, new()));
        Assert.Equal(3, handler.Query("bookmark").Count);
        Assert.Empty(handler.Validate());
    }
}
