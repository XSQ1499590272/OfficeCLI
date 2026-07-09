using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordTabContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddSetAndRemoveParagraphTabStop_ReadsBackCurrentTabCollection()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "tab host" });
        var tabPath = handler.Add(paragraphPath, "tab", null, new()
        {
            ["pos"] = "1440",
            ["val"] = "right",
            ["leader"] = "dot"
        });

        handler.Set(tabPath, new()
        {
            ["pos"] = "720",
            ["val"] = "center",
            ["leader"] = "hyphen"
        });
        var paragraph = handler.Get(paragraphPath);

        var tab = Assert.Single((IEnumerable<Dictionary<string, object?>>)Fmt(paragraph)["tabs"]!);
        Assert.Equal(720, Convert.ToInt32(tab["pos"]));
        Assert.Equal("center", tab["val"]);
        Assert.Equal("hyphen", tab["leader"]);

        handler.Remove(tabPath);
        var updated = handler.Get(paragraphPath);

        Assert.False(Fmt(updated).ContainsKey("tabs"));
        var error = Assert.Single(handler.Validate());
        Assert.Equal("Schema", error.ErrorType);
        Assert.Contains("incomplete content", error.Description);
    }

    [Fact]
    public void AddAndSetPositionalTab_ReadsBackInlineRunProperties()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var ptabPath = handler.Add(paragraphPath, "ptab", null, new()
        {
            ["align"] = "right",
            ["relativeTo"] = "margin",
            ["leader"] = "dot"
        });

        var ptab = handler.Get(ptabPath);
        var queried = Assert.Single(handler.Query("ptab"));

        Assert.Equal("ptab", ptab.Type);
        Assert.Equal("right", Fmt(ptab)["align"]);
        Assert.Equal("margin", Fmt(ptab)["relativeTo"]);
        Assert.Equal("dot", Fmt(ptab)["leader"]);
        Assert.Equal(ptabPath, queried.Path);

        handler.Set(ptabPath, new()
        {
            ["align"] = "center",
            ["leader"] = "underscore"
        });
        var updated = handler.Get(ptabPath);

        Assert.Equal("center", Fmt(updated)["align"]);
        Assert.Equal("underscore", Fmt(updated)["leader"]);
        Assert.Empty(handler.Validate());
    }
}
