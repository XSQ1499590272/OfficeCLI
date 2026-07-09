using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordStyleAndNumberingTests : WordTestBase
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

        Assert.Equal("/styles/MyHeading", stylePath);
        Assert.Equal("style", style.Type);
        Assert.Equal("MyHeading", Fmt(style)["id"]);
        Assert.Equal("My Heading", Fmt(style)["name"]);
        Assert.Equal("paragraph", Fmt(style)["type"]);
        Assert.Equal("Normal", Fmt(style)["basedOn"]);
        Assert.Equal("MyHeading", Fmt(paragraph)["styleId"]);
        Assert.Equal("My Heading", Fmt(paragraph)["styleName"]);
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

    [Fact]
    public void AddParagraph_StyleNameResolvesExistingNormalAndSkipsMissingSpacedName()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var normalPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "normal",
            ["styleName"] = "Normal"
        });
        var missingPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "missing heading",
            ["styleName"] = "Heading 1"
        });

        var normal = handler.Get(normalPath);
        var missing = handler.Get(missingPath);

        Assert.Equal("Normal", Fmt(normal)["styleId"]);
        Assert.Equal("Normal", Fmt(normal)["styleName"]);
        Assert.False(Fmt(missing).ContainsKey("styleId"));
        Assert.False(Fmt(missing).ContainsKey("styleName"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddCharacterAndTableStyles_ReadsBackTypeNameAndBasedOn()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var characterPath = handler.Add("/styles", "style", null, new()
        {
            ["id"] = "EmphasisChar",
            ["name"] = "Emphasis Character",
            ["type"] = "character",
            ["basedOn"] = "DefaultParagraphFont",
            ["bold"] = "true"
        });
        var tablePath = handler.Add("/styles", "style", null, new()
        {
            ["id"] = "CompactTable",
            ["name"] = "Compact Table",
            ["type"] = "table",
            ["basedOn"] = "TableNormal"
        });

        var character = handler.Get(characterPath);
        var table = handler.Get(tablePath);

        Assert.Equal("/styles/EmphasisChar", characterPath);
        Assert.Equal("EmphasisChar", Fmt(character)["id"]);
        Assert.Equal("Emphasis Character", Fmt(character)["name"]);
        Assert.Equal("character", Fmt(character)["type"]);
        Assert.Equal("DefaultParagraphFont", Fmt(character)["basedOn"]);
        Assert.Equal(true, Fmt(character)["bold"]);

        Assert.Equal("/styles/CompactTable", tablePath);
        Assert.Equal("CompactTable", Fmt(table)["id"]);
        Assert.Equal("Compact Table", Fmt(table)["name"]);
        Assert.Equal("table", Fmt(table)["type"]);
        Assert.Equal("TableNormal", Fmt(table)["basedOn"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddAbstractNumNumAndParagraphListStyle_ReadsBackCanonicalNumbering()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var abstractPath = handler.Add("/numbering", "abstractNum", null, new()
        {
            ["id"] = "7",
            ["type"] = "multilevel",
            ["level0.format"] = "decimal",
            ["level0.text"] = "%1.",
            ["level0.indent"] = "720"
        });
        var numPath = handler.Add("/numbering", "num", null, new()
        {
            ["id"] = "3",
            ["abstractNumId"] = "7",
            ["start"] = "4"
        });
        var paragraphPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "ordered item",
            ["listStyle"] = "ordered",
            ["start"] = "3",
            ["numLevel"] = "1"
        });

        var abstractNum = handler.Get(abstractPath, depth: 1);
        var level = handler.Get($"{abstractPath}/level[0]");
        var num = handler.Get(numPath);
        var paragraph = handler.Get(paragraphPath);

        Assert.Equal("7", Fmt(abstractNum)["id"]);
        Assert.Equal("multilevel", Fmt(abstractNum)["type"]);
        Assert.Equal("decimal", Fmt(level)["format"]);
        Assert.Equal("%1.", Fmt(level)["lvlText"]);
        Assert.Equal("720", Fmt(level)["indent"]);
        Assert.Equal("7", Fmt(num)["abstractNumId"]);
        Assert.Equal("4", Fmt(num)["startOverride.0"]);
        Assert.Equal("ordered", Fmt(paragraph)["listStyle"]);
        Assert.Equal("lowerLetter", Fmt(paragraph)["numFmt"]);
        Assert.Equal("1", Fmt(paragraph)["numLevel"]);

        handler.Set(paragraphPath, new() { ["listStyle"] = "none" });
        Assert.False(Fmt(handler.Get(paragraphPath)).ContainsKey("listStyle"));
        Assert.Empty(handler.Validate());
    }
}
