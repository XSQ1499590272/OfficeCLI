using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordStyleContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddStyle_ReadsBackStyleAndParagraphStyleReference()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var stylePath = handler.Add("/styles", "style", null, new()
        {
            ["id"] = "MyHeading",
            ["name"] = "My Heading",
            ["type"] = "paragraph",
            ["basedOn"] = "Normal",
            ["align"] = "center",
            ["bold"] = "true"
        });
        var paragraphPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Styled paragraph",
            ["style"] = "MyHeading"
        });

        var style = handler.Get(stylePath);
        var paragraph = handler.Get(paragraphPath);
        var queried = Assert.Single(handler.Query("style:contains(\"My Heading\")"));

        Assert.Equal("/styles/MyHeading", stylePath);
        Assert.Equal("style", style.Type);
        Assert.Equal("MyHeading", Fmt(style)["id"]);
        Assert.Equal("My Heading", Fmt(style)["name"]);
        Assert.Equal("paragraph", Fmt(style)["type"]);
        Assert.Equal("Normal", Fmt(style)["basedOn"]);
        Assert.Equal(stylePath, queried.Path);
        Assert.Equal("MyHeading", Fmt(paragraph)["styleId"]);
        Assert.Equal("My Heading", Fmt(paragraph)["styleName"]);
    }

    [Fact]
    public void AddParagraph_WithDisplayNameStyle_WritesValueAsStyleId()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/styles", "style", null, new()
        {
            ["id"] = "DisplayNameStyle",
            ["name"] = "Display Name Style",
            ["type"] = "paragraph"
        });

        var paragraphPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "display lookup",
            ["style"] = "Display Name Style"
        });
        var paragraph = handler.Get(paragraphPath);

        Assert.Equal("Display Name Style", Fmt(paragraph)["styleId"]);
        Assert.Equal("Display Name Style", Fmt(paragraph)["styleName"]);
    }

    [Fact]
    public void FailedAdd_DuplicateCustomStyleIdDoesNotMutateStyles()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/styles", "style", null, new()
        {
            ["id"] = "UniqueStyle",
            ["name"] = "Unique Style",
            ["type"] = "paragraph"
        });
        var before = handler.Query("style").Count;

        var ex = Assert.Throws<ArgumentException>(() =>
            handler.Add("/styles", "style", null, new()
            {
                ["id"] = "UniqueStyle",
                ["name"] = "Another Name",
                ["type"] = "paragraph"
            }));

        Assert.Contains("already exists", ex.Message);
        Assert.Equal(before, handler.Query("style").Count);
    }
}
