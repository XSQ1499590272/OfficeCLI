using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordTableTests : WordTestBase
{
    [Fact]
    public void AddTable_CreatesRowsColumnsCellsAndCanonicalLayout()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["data"] = "A,B;C,D",
            ["colWidths"] = "1200,1800",
            ["layout"] = "fixed"
        });

        var table = handler.Get(tablePath, depth: 2);

        Assert.Equal("table", table.Type);
        Assert.Equal("fixed", Fmt(table)["layout"]);
        Assert.Equal("1200dxa,1800dxa", Fmt(table)["colWidths"]);
        Assert.Equal(2, table.Children.Count);
        Assert.All(table.Children, row => Assert.Equal(2, row.Children.Count));
        Assert.Equal("A", table.Children[0].Children[0].Text);
        Assert.Equal("D", table.Children[1].Children[1].Text);
    }

    [Fact]
    public void SetCellFormattingAndMerges_ReadsBackCurrentIndexBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new()
        {
            ["data"] = "A,B,C;D,E,F",
            ["colWidths"] = "1000,1000,1000"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["fill"] = "F1FAEE",
            ["align"] = "center",
            ["valign"] = "center",
            ["padding"] = "120",
            ["gridspan"] = "2"
        });
        handler.Set("/body/tbl[1]/tr[2]/tc[1]", new() { ["vmerge"] = "continue" });

        var topRow = handler.Get("/body/tbl[1]/tr[1]", depth: 1);
        var mergedCell = topRow.Children[0];
        var shiftedCell = handler.Get("/body/tbl[1]/tr[1]/tc[2]");
        var vmergeCell = handler.Get("/body/tbl[1]/tr[2]/tc[1]");

        Assert.Equal(2, topRow.Children.Count);
        Assert.Equal(2, Fmt(mergedCell)["colspan"]);
        Assert.Equal("#F1FAEE", Fmt(mergedCell)["fill"]);
        Assert.Equal("center", Fmt(mergedCell)["align"]);
        Assert.Equal("center", Fmt(mergedCell)["valign"]);
        Assert.Equal(120, Fmt(mergedCell)["padding.top"]);
        Assert.Equal("C", shiftedCell.Text);
        Assert.Equal("continue", Fmt(vmergeCell)["vmerge"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddAndRemoveVirtualColumn_UpdatesGridAndRows()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["data"] = "A,B;C,D",
            ["colWidths"] = "1000,1000"
        });

        var columnPath = handler.Add(tablePath, "column", null, new()
        {
            ["width"] = "1800",
            ["text"] = "X"
        });

        Assert.Equal("/body/tbl[1]/col[3]", columnPath);
        Assert.Equal("1000dxa,1000dxa,1800dxa", Fmt(handler.Get(tablePath))["colWidths"]);

        handler.Remove("/body/tbl[1]/col[2]");
        var table = handler.Get(tablePath, depth: 2);

        Assert.Equal("1000dxa,1800dxa", Fmt(table)["colWidths"]);
        Assert.All(table.Children, row => Assert.Equal(2, row.Children.Count));
        Assert.Equal("A", table.Children[0].Children[0].Text);
        Assert.Equal("X", table.Children[0].Children[1].Text);
        Assert.Empty(handler.Validate());
    }
}
