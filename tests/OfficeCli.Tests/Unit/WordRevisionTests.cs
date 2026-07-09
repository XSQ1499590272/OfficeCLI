using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordRevisionTests : WordTestBase
{
    [Fact]
    public void ParagraphRevisionFormatAndUnsupportedInsertion_ReadBackCurrentBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "paragraph text" });

        handler.Set(paragraphPath, new()
        {
            ["revision.type"] = "format",
            ["revision.author"] = "Dana",
            ["revision.id"] = "201",
            ["align"] = "center"
        });
        var revision = Assert.Single(handler.Query("revision[@id=201]"));
        var paragraph = handler.Get(paragraphPath);

        Assert.Equal("/revision[@id=201]", revision.Path);
        Assert.Equal("paragraph", Fmt(revision)["revision.type"]);
        Assert.Equal("Dana", Fmt(revision)["revision.author"]);
        Assert.Equal("201", Fmt(revision)["revision.id"]);
        Assert.StartsWith("/body/p", Fmt(revision)["revision.nativePath"]!.ToString());
        Assert.Equal("center", Fmt(paragraph)["align"]);
        Assert.Equal("format", Fmt(paragraph)["revision.type"]);

        var ex = Assert.Throws<InvalidOperationException>(() =>
            handler.Set(paragraphPath, new()
            {
                ["revision.type"] = "ins",
                ["revision.author"] = "Eve",
                ["revision.id"] = "202"
            }));

        Assert.Contains("not supported via set", ex.Message);
        Assert.Empty(handler.Query("revision[@id=202]"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void TableRowAndCellRevisionMarkers_ReadBackCurrentStructuralTypes()
    {
        var rowPath = CreateBlankDocx();

        using var rowHandler = new WordHandler(rowPath, editable: true);
        rowHandler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        rowHandler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["revision.type"] = "ins",
            ["revision.author"] = "Row Author",
            ["revision.id"] = "401"
        });
        var rowRevision = Assert.Single(rowHandler.Query("revision[@id=401]"));
        var row = rowHandler.Get("/body/tbl[1]/tr[1]");

        Assert.Equal("rowIns", Fmt(rowRevision)["revision.type"]);
        Assert.Equal("Row Author", Fmt(rowRevision)["revision.author"]);
        Assert.Equal("ins", Fmt(row)["revision.type"]);
        Assert.Equal("Row Author", Fmt(row)["revision.author"]);
        Assert.Empty(rowHandler.Validate());
    }

    [Fact]
    public void RowRevisionFormatMarker_ReadsBackTrPrChange()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        handler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["revision.type"] = "format",
            ["revision.author"] = "Row Format",
            ["revision.id"] = "421",
            ["height.exact"] = "500"
        });

        var revision = Assert.Single(handler.Query("revision[@id=421]"));
        var row = handler.Get("/body/tbl[1]/tr[1]");

        Assert.Equal("format", Fmt(revision)["revision.type"]);
        Assert.Equal("Row Format", Fmt(revision)["revision.author"]);
        Assert.Equal("/body/tbl[1]/tr[1]", Fmt(revision)["revision.nativePath"]);
        Assert.Equal("Row Format", Fmt(row)["trPrChange.author"]);
        Assert.Equal("500dxa", Fmt(row)["height"]);
        Assert.Equal("exact", Fmt(row)["height.rule"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void TableAndCellRevisionFormatMarkers_ReadBackCurrentStructuralTypes()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        handler.Set("/body/tbl[1]", new()
        {
            ["revision.type"] = "format",
            ["revision.author"] = "Table Author",
            ["revision.id"] = "411",
            ["layout"] = "fixed"
        });
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["revision.type"] = "format",
            ["revision.author"] = "Cell Author",
            ["revision.id"] = "412",
            ["fill"] = "F1FAEE"
        });

        var tableRevision = Assert.Single(handler.Query("revision[@id=411]"));
        var cellRevision = Assert.Single(handler.Query("revision[@id=412]"));
        var table = handler.Get("/body/tbl[1]");
        var cell = handler.Get("/body/tbl[1]/tr[1]/tc[1]");

        Assert.Equal("format", Fmt(tableRevision)["revision.type"]);
        Assert.Equal("Table Author", Fmt(tableRevision)["revision.author"]);
        Assert.Equal("format", Fmt(cellRevision)["revision.type"]);
        Assert.Equal("Cell Author", Fmt(cellRevision)["revision.author"]);
        Assert.Equal("fixed", Fmt(table)["layout"]);
        Assert.Equal("#F1FAEE", Fmt(cell)["fill"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void MoveRunWithRevisionProps_CreatesPairedMoveFromMoveToRevisions()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var sourcePara = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var movedRun = handler.Add(sourcePara, "run", null, new() { ["text"] = "Moved text" });
        var targetPara = handler.Add("/body", "paragraph", null, new() { ["text"] = "Target: " });

        var movedPath = handler.Move(movedRun, targetPara, null, new()
        {
            ["revision.author"] = "Reviewer",
            ["revision.date"] = "2026-07-09T01:02:03Z",
            ["revision.id"] = "77"
        });

        var revisions = handler.Query("revision");
        var moveFrom = Assert.Single(revisions, r =>
            Fmt(r).GetValueOrDefault("revision.type")?.ToString() == "moveFrom");
        var moveTo = Assert.Single(revisions, r =>
            Fmt(r).GetValueOrDefault("revision.type")?.ToString() == "moveTo");

        Assert.Equal("Moved text", handler.Get(movedPath).Text);
        Assert.Equal("Moved text", moveFrom.Text);
        Assert.Equal("Moved text", moveTo.Text);
        Assert.Equal("77", Fmt(moveFrom)["revision.id"]);
        Assert.Equal("77", Fmt(moveTo)["revision.id"]);
        Assert.Equal("Reviewer", Fmt(moveFrom)["revision.author"]);
        Assert.Equal("Reviewer", Fmt(moveTo)["revision.author"]);
        Assert.EndsWith("[@type=moveFrom]", moveFrom.Path);
        Assert.EndsWith("[@type=moveTo]", moveTo.Path);
        Assert.Empty(handler.Validate());
    }
}
