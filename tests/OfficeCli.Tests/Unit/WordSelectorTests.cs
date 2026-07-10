using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordSelectorTests : WordTestBase
{
    [Fact]
    public void QuerySelector_SupportsChildCombinatorContainsEmptyAndCaseInsensitiveElement()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "The quick fox",
            ["align"] = "center"
        });
        handler.Add(first, "run", null, new()
        {
            ["text"] = " jumps",
            ["bold"] = "true"
        });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "plain text" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "" });

        var boldChildren = handler.Query("PARAGRAPH[align=center] > RUN[bold=true]");
        var contains = handler.Query("paragraph:contains(\"quick fox\")");
        var empty = handler.Query("paragraph:empty");

        Assert.Single(boldChildren);
        Assert.Equal("run", boldChildren[0].Type);
        Assert.Equal(" jumps", boldChildren[0].Text);
        Assert.Single(contains);
        Assert.Equal(first, contains[0].Path);
        Assert.Single(empty);
        Assert.Equal("", empty[0].Text);
    }

    [Fact]
    public void QuerySelector_FixesCurrentDirectParagraphFilterBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Alpha needle",
            ["align"] = "center"
        });
        var second = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Beta haystack",
            ["align"] = "right"
        });

        var rightAligned = handler.Query("paragraph[align=right]");
        var notCenter = handler.Query("paragraph[align!=center]");
        var textContains = handler.Query("paragraph[text~=needle]");
        var typeFallback = handler.Query("paragraph[type=paragraph]");

        var aligned = Assert.Single(rightAligned);
        Assert.Equal(second, aligned.Path);
        var notCenterNode = Assert.Single(notCenter);
        Assert.Equal(second, notCenterNode.Path);
        Assert.Equal(2, textContains.Count);
        Assert.Contains(textContains, node => node.Path == first);
        Assert.Contains(textContains, node => node.Path == second);
        Assert.Equal(2, typeFallback.Count);
        Assert.All(typeFallback, node => Assert.Equal("paragraph", node.Type));
    }

    [Fact]
    public void QueryAndGet_RejectUnclosedSelectorOrPathInsteadOfReturningEmpty()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "parse boundary" });

        var selectorError = Assert.Throws<ArgumentException>(
            () => handler.Query("paragraph[align=center"));
        Assert.Contains("unclosed bracket", selectorError.Message);

        var pathError = Assert.Throws<ArgumentException>(
            () => handler.Get("/body/p[1"));
        Assert.Contains("Malformed", pathError.Message);

        Assert.Single(handler.Query("paragraph"));
    }
}
