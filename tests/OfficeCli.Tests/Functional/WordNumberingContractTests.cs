using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordNumberingContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddAbstractNumAndNum_ReadsBackCanonicalNumberingPaths()
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

        var numbering = handler.Get("/numbering", depth: 1);
        var abstractNum = handler.Get(abstractPath, depth: 1);
        var level = handler.Get($"{abstractPath}/level[0]");
        var num = handler.Get(numPath);

        Assert.Equal("numbering", numbering.Type);
        Assert.Contains(numbering.Children, child => child.Path == "/numbering/abstractNum[1]" && child.Type == "abstractNum");
        Assert.Contains(numbering.Children, child => child.Path == "/numbering/num[1]" && child.Type == "num");
        Assert.Equal("7", Fmt(abstractNum)["id"]);
        Assert.Equal("multilevel", Fmt(abstractNum)["type"]);
        Assert.Equal("decimal", Fmt(level)["format"]);
        Assert.Equal("%1.", Fmt(level)["lvlText"]);
        Assert.Equal("720", Fmt(level)["indent"]);
        Assert.Equal("7", Fmt(num)["abstractNumId"]);
        Assert.Equal("4", Fmt(num)["startOverride.0"]);
    }

    [Fact]
    public void ParagraphListStyle_AddsAndClearsHighLevelNumbering()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var orderedPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "ordered item",
            ["listStyle"] = "ordered",
            ["start"] = "3",
            ["numLevel"] = "1"
        });
        var bulletPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "bullet item",
            ["listStyle"] = "bullet"
        });

        var ordered = handler.Get(orderedPath);
        var bullet = handler.Get(bulletPath);

        Assert.Equal("ordered", Fmt(ordered)["listStyle"]);
        Assert.Equal("lowerLetter", Fmt(ordered)["numFmt"]);
        Assert.Equal("1", Fmt(ordered)["numLevel"]);
        Assert.Equal("bullet", Fmt(bullet)["listStyle"]);
        Assert.Equal("bullet", Fmt(bullet)["numFmt"]);

        handler.Set(orderedPath, new() { ["listStyle"] = "none" });
        var cleared = handler.Get(orderedPath);

        Assert.False(Fmt(cleared).ContainsKey("listStyle"));
        Assert.Empty(handler.Validate());
    }
}
