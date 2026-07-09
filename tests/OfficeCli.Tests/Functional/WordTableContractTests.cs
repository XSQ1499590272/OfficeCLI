using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordTableContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddTable_CreatesRowsColumnsAndReadableCells()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["data"] = "A,B;C,D",
            ["colWidths"] = "1200,1800"
        });

        var table = handler.Get(tablePath, depth: 2);

        Assert.Equal("table", table.Type);
        Assert.Equal("1200dxa,1800dxa", Fmt(table)["colWidths"]);
        Assert.Equal(2, table.Children.Count);
        Assert.All(table.Children, row => Assert.Equal(2, row.Children.Count));
        Assert.Equal("A", table.Children[0].Children[0].Text);
        Assert.Equal("D", table.Children[1].Children[1].Text);
    }

    [Fact]
    public void AddAndSetTableLayoutBordersAndSpacing_ReadsBackCanonicalProperties()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["data"] = "A,B",
            ["layout"] = "fixed",
            ["indent"] = "240",
            ["cellSpacing"] = "120",
            ["border"] = "none",
            ["border.top"] = "double;8;FF0000"
        });

        handler.Set(tablePath, new()
        {
            ["colWidths"] = "900,1500",
            ["border.top.color"] = "00AA00",
            ["border.insideH"] = "single;4;0000FF"
        });
        var table = handler.Get(tablePath);

        Assert.Equal("fixed", Fmt(table)["layout"]);
        Assert.Equal(240, Convert.ToInt32(Fmt(table)["indent"]));
        Assert.Equal(120, Convert.ToInt32(Fmt(table)["cellSpacing"]));
        Assert.Equal("900dxa,1500dxa", Fmt(table)["colWidths"]);
        Assert.Equal("double", Fmt(table)["border.top"]);
        Assert.Equal("#00AA00", Fmt(table)["border.top.color"]);
        Assert.Equal("single", Fmt(table)["border.insideH"]);
        Assert.Equal("#0000FF", Fmt(table)["border.insideH.color"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddCell_ReadsBackTextWidthFillAlignmentValignAndPadding()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        var cellPath = handler.Add("/body/tbl[1]/tr[1]", "cell", null, new()
        {
            ["text"] = "Added cell",
            ["width"] = "1440",
            ["fill"] = "F1FAEE",
            ["align"] = "center",
            ["valign"] = "center",
            ["padding"] = "120"
        });

        var cell = handler.Get(cellPath);

        Assert.Equal("cell", cell.Type);
        Assert.Equal("Added cell", cell.Text);
        Assert.Equal("1440dxa", Fmt(cell)["width"]);
        Assert.Equal("#F1FAEE", Fmt(cell)["fill"]);
        Assert.Equal("center", Fmt(cell)["align"]);
        Assert.Equal("center", Fmt(cell)["valign"]);
        Assert.Equal(120, Fmt(cell)["padding.top"]);
        Assert.Equal(120, Fmt(cell)["padding.bottom"]);
        Assert.Equal(120, Fmt(cell)["padding.left"]);
        Assert.Equal(120, Fmt(cell)["padding.right"]);
    }

    [Fact]
    public void SetCellGridSpan_RemovesAbsorbedCellAndShiftsFollowingCellIndex()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new()
        {
            ["data"] = "A,B,C",
            ["colWidths"] = "1000,1000,1000"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["gridspan"] = "2" });
        var row = handler.Get("/body/tbl[1]/tr[1]", depth: 1);
        var shiftedCell = handler.Get("/body/tbl[1]/tr[1]/tc[2]");

        Assert.Equal(2, row.Children.Count);
        Assert.Equal("A", row.Children[0].Text);
        Assert.Equal(2, Fmt(row.Children[0])["colspan"]);
        Assert.Equal("C", shiftedCell.Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetCellVerticalMerge_ReadsBackRestartAndContinue()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new()
        {
            ["data"] = "Top,Right;Bottom,Other",
            ["colWidths"] = "1200,1200"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["vmerge"] = "restart" });
        handler.Set("/body/tbl[1]/tr[2]/tc[1]", new() { ["vmerge"] = "continue" });

        var top = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        var bottom = handler.Get("/body/tbl[1]/tr[2]/tc[1]");

        Assert.Equal("restart", Fmt(top)["vmerge"]);
        Assert.Equal("continue", Fmt(bottom)["vmerge"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetCellHMergeRestart_UsesGridSpanSemantics()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new()
        {
            ["data"] = "A,B,C",
            ["colWidths"] = "1000,1000,1000"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["hmerge"] = "restart" });
        var row = handler.Get("/body/tbl[1]/tr[1]", depth: 1);

        Assert.Equal(2, row.Children.Count);
        Assert.Equal("A", row.Children[0].Text);
        Assert.Equal(2, Fmt(row.Children[0])["colspan"]);
        Assert.Equal("C", row.Children[1].Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddAndRemoveVirtualColumn_UpdatesGridAndEveryRow()
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
        var expanded = handler.Get(tablePath, depth: 2);

        Assert.Equal("/body/tbl[1]/col[3]", columnPath);
        Assert.Equal("1000dxa,1000dxa,1800dxa", Fmt(expanded)["colWidths"]);
        Assert.All(expanded.Children, row => Assert.Equal(3, row.Children.Count));
        Assert.Equal("X", expanded.Children[0].Children[2].Text);
        Assert.Equal("X", expanded.Children[1].Children[2].Text);

        handler.Remove("/body/tbl[1]/col[2]");
        var shrunk = handler.Get(tablePath, depth: 2);

        Assert.Equal("1000dxa,1800dxa", Fmt(shrunk)["colWidths"]);
        Assert.All(shrunk.Children, row => Assert.Equal(2, row.Children.Count));
        Assert.Equal("A", shrunk.Children[0].Children[0].Text);
        Assert.Equal("X", shrunk.Children[0].Children[1].Text);
        Assert.Equal("C", shrunk.Children[1].Children[0].Text);
        Assert.Equal("X", shrunk.Children[1].Children[1].Text);
        Assert.Empty(handler.Validate());
    }
}
