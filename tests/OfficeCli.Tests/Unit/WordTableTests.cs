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

    [Fact]
    public void SetRowAndCellPropertyMatrix_ReadsBackCurrentCanonicalValues()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["data"] = "A,B" });

        handler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["height.atleast"] = "720",
            ["header"] = "true",
            ["cantSplit"] = "true",
            ["hidden"] = "true",
            ["cellSpacing"] = "40",
            ["rowAlign"] = "right",
            ["gridBefore"] = "1",
            ["wBefore"] = "240"
        });
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["width"] = "50%",
            ["fill"] = "FFF2CC",
            ["tcFitText"] = "true",
            ["nowrap"] = "true",
            ["textDirection"] = "tbRlV",
            ["padding.top"] = "60",
            ["padding.bottom"] = "80",
            ["padding.left"] = "100",
            ["padding.right"] = "120"
        });

        var row = handler.Get("/body/tbl[1]/tr[1]");
        var cell = handler.Get("/body/tbl[1]/tr[1]/tc[1]");

        Assert.Equal("720dxa", Fmt(row)["height"]);
        Assert.Equal("atLeast", Fmt(row)["height.rule"]);
        Assert.Equal(true, Fmt(row)["header"]);
        Assert.Equal(true, Fmt(row)["cantSplit"]);
        Assert.Equal(true, Fmt(row)["hidden"]);
        Assert.Equal(40, Convert.ToInt32(Fmt(row)["cellSpacing"]));
        Assert.Equal("right", Fmt(row)["rowAlign"]);
        Assert.Equal("1", Fmt(row)["gridBefore"]);
        Assert.Equal("240dxa", Fmt(row)["wBefore"]);
        Assert.Equal("50%", Fmt(cell)["width"]);
        Assert.Equal("#FFF2CC", Fmt(cell)["fill"]);
        Assert.Equal(true, Fmt(cell)["tcFitText"]);
        Assert.Equal(true, Fmt(cell)["nowrap"]);
        Assert.Equal("tbRlV", Fmt(cell)["textDirection"]);
        Assert.Equal(60, Fmt(cell)["padding.top"]);
        Assert.Equal(80, Fmt(cell)["padding.bottom"]);
        Assert.Equal(100, Fmt(cell)["padding.left"]);
        Assert.Equal(120, Fmt(cell)["padding.right"]);
        Assert.Empty(handler.Validate());
    }
}
