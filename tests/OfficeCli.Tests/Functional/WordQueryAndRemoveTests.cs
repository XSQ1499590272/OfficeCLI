using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordQueryAndRemoveTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void Query_ReturnsStablePathsAndEmptyResults()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var targetPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "needle paragraph" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "plain paragraph" });

        var matches = handler.Query("paragraph:contains(\"needle\")");
        var misses = handler.Query("paragraph:contains(\"missing\")");

        Assert.Single(matches);
        Assert.Equal(targetPath, matches[0].Path);
        Assert.Empty(misses);
    }

    [Fact]
    public void Remove_DeletesParagraphWithoutDeletingSiblings()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var removePath = handler.Add("/body", "paragraph", null, new() { ["text"] = "remove me" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "keep me" });

        handler.Remove(removePath);

        Assert.Empty(handler.Query("paragraph:contains(\"remove me\")"));
        var kept = Assert.Single(handler.Query("paragraph:contains(\"keep me\")"));
        Assert.Equal("keep me", kept.Text);
    }
}
