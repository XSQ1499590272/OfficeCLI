using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordTableRowContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddRow_ReadsBackCellsHeightAndRowProperties()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var tablePath = handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });
        var rowPath = handler.Add(tablePath, "row", null, new()
        {
            ["cols"] = "3",
            ["c1"] = "A",
            ["c2"] = "B",
            ["c3"] = "C",
            ["height.exact"] = "720",
            ["header"] = "true",
            ["cantSplit"] = "true",
            ["rowAlign"] = "center"
        });

        var table = handler.Get(tablePath, depth: 2);
        var row = handler.Get(rowPath, depth: 1);

        Assert.Equal("row", row.Type);
        Assert.Equal("720dxa", Fmt(row)["height"]);
        Assert.Equal("exact", Fmt(row)["height.rule"]);
        Assert.Equal(true, Fmt(row)["header"]);
        Assert.Equal(true, Fmt(row)["cantSplit"]);
        Assert.Equal("center", Fmt(row)["rowAlign"]);
        Assert.Equal(2, table.Children.Count);
        Assert.All(table.Children, child => Assert.Equal(3, child.Children.Count));
        Assert.Equal("A", row.Children[0].Text);
        Assert.Equal("B", row.Children[1].Text);
        Assert.Equal("C", row.Children[2].Text);
    }

    [Fact]
    public void SetRow_UpdatesCellTextAndClearsBooleanProperties()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });
        handler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["header"] = "true",
            ["cantSplit"] = "true",
            ["c1"] = "before"
        });

        handler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["header"] = "false",
            ["cantSplit"] = "false",
            ["c1"] = "after",
            ["c2"] = "second"
        });
        var row = handler.Get("/body/tbl[1]/tr[1]", depth: 1);

        Assert.False(Fmt(row).ContainsKey("header"));
        Assert.False(Fmt(row).ContainsKey("cantSplit"));
        Assert.Equal("after", row.Children[0].Text);
        Assert.Equal("second", row.Children[1].Text);
    }
}
