using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordBodyMutationTests : WordTestBase
{
    [Fact]
    public void AddParagraphQueryAndRemove_PreserveStableBodySiblings()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var removePath = handler.Add("/body", "paragraph", null, new() { ["text"] = "remove me" });
        var keepPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "keep me" });

        var match = Assert.Single(handler.Query("paragraph:contains(\"remove\")"));
        Assert.Equal(removePath, match.Path);
        Assert.Contains("@paraId=", keepPath);

        handler.Remove(removePath);

        Assert.Empty(handler.Query("paragraph:contains(\"remove me\")"));
        var kept = Assert.Single(handler.Query("paragraph:contains(\"keep me\")"));
        Assert.Equal("keep me", kept.Text);
        Assert.Equal(keepPath, kept.Path);
    }

    [Fact]
    public void MoveAndCopy_ReorderAndCloneBodyParagraphsWithoutLosingText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.Add("/body", "paragraph", null, new() { ["text"] = "first" });
        var second = handler.Add("/body", "paragraph", null, new() { ["text"] = "second" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "third" });

        handler.Move(first, "/body", InsertPosition.AfterElement(second));
        var movedParagraphs = handler.Get("/body", depth: 1)
            .Children.Where(child => child.Type == "paragraph").Select(child => child.Text).ToList();

        Assert.Equal(new[] { "second", "first", "third" }, movedParagraphs);

        var copyPath = handler.CopyFrom(second, "/body", null);
        var copies = handler.Query("paragraph:contains(\"second\")");

        Assert.Equal("second", handler.Get(copyPath).Text);
        Assert.Equal(2, copies.Count);
        Assert.Empty(handler.Validate());
    }
}
