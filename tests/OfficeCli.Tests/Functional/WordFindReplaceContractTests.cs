using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordFindReplaceContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void SetFindReplace_ReplacesOnlyMatchingParagraphs()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.Add("/body", "paragraph", null, new() { ["text"] = "needle here" });
        var second = handler.Add("/body", "paragraph", null, new() { ["text"] = "plain here" });

        var unsupported = handler.Set("/body", new()
        {
            ["find"] = "needle",
            ["replace"] = "thread"
        });

        Assert.Empty(unsupported);
        Assert.Equal(1, handler.LastFindMatchCount);
        Assert.Equal("thread here", handler.Get(first).Text);
        Assert.Equal("plain here", handler.Get(second).Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetRegexFindReplace_ExpandsRegexBackreferences()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Invoice 42" });

        handler.Set(paragraphPath, new()
        {
            ["find"] = "(\\d+)",
            ["regex"] = "true",
            ["replace"] = "#$1"
        });

        Assert.Equal("Invoice #42", handler.Get(paragraphPath).Text);
        Assert.Equal(1, handler.LastFindMatchCount);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedSet_BareFindWithoutActionThrows()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "needle" });

        var ex = Assert.Throws<ArgumentException>(() =>
            handler.Set("/body", new() { ["find"] = "needle" }));

        Assert.Contains("'find' requires", ex.Message);
    }
}
