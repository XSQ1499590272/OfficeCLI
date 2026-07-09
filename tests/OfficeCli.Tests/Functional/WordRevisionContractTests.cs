using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordRevisionContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void SetRunRevisionInsertion_QueryReturnsSyntheticRevisionNode()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paragraphPath, "run", null, new() { ["text"] = "inserted text" });

        handler.Set(runPath, new()
        {
            ["revision.type"] = "ins",
            ["revision.author"] = "Alice",
            ["revision.date"] = "2026-01-02T03:04:05Z",
            ["revision.id"] = "101"
        });
        var revision = Assert.Single(handler.Query("revision"));

        Assert.Equal("/revision[@id=101]", revision.Path);
        Assert.Equal("revision", revision.Type);
        Assert.Equal("inserted text", revision.Text);
        Assert.Equal("ins", Fmt(revision)["revision.type"]);
        Assert.Equal("Alice", Fmt(revision)["revision.author"]);
        Assert.Equal("101", Fmt(revision)["revision.id"]);
        Assert.StartsWith("/body/p", Fmt(revision)["revision.nativePath"]!.ToString());
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetRunRevisionDeletion_QueryReturnsDeletedText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paragraphPath, "run", null, new() { ["text"] = "deleted text" });

        handler.Set(runPath, new()
        {
            ["revision.type"] = "del",
            ["revision.author"] = "Bob",
            ["revision.id"] = "102"
        });
        var revision = Assert.Single(handler.Query("revision[@type=del]"));

        Assert.Equal("deleted text", revision.Text);
        Assert.Equal("del", Fmt(revision)["revision.type"]);
        Assert.Equal("Bob", Fmt(revision)["revision.author"]);
        Assert.Equal("102", Fmt(revision)["revision.id"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetRunRevisionFormat_QueryReturnsFormatMarker()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paragraphPath, "run", null, new() { ["text"] = "formatted text" });

        handler.Set(runPath, new()
        {
            ["revision.type"] = "format",
            ["revision.author"] = "Carol",
            ["revision.id"] = "103",
            ["bold"] = "true"
        });
        var revision = Assert.Single(handler.Query("revision[@type=format]"));
        var run = handler.Get(runPath);

        Assert.Equal("format", Fmt(revision)["revision.type"]);
        Assert.Equal("Carol", Fmt(revision)["revision.author"]);
        Assert.Equal("103", Fmt(revision)["revision.id"]);
        Assert.Equal(true, Fmt(run)["bold"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetParagraphRevisionFormat_QueryReturnsParagraphPropertyMarker()
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
        var revision = Assert.Single(handler.Query("revision[@type=format]"));
        var paragraph = handler.Get(paragraphPath);

        Assert.Equal("/revision[@id=201]", revision.Path);
        Assert.Equal("paragraph", Fmt(revision)["revision.type"]);
        Assert.Equal("Dana", Fmt(revision)["revision.author"]);
        Assert.Equal("201", Fmt(revision)["revision.id"]);
        Assert.StartsWith("/body/p", Fmt(revision)["revision.nativePath"]!.ToString());
        Assert.Equal("center", Fmt(paragraph)["align"]);
        Assert.Equal("format", Fmt(paragraph)["revision.type"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetParagraphRevisionInsertionIsCurrentlyRejected()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "paragraph text" });

        var ex = Assert.Throws<InvalidOperationException>(() =>
            handler.Set(paragraphPath, new()
            {
                ["revision.type"] = "ins",
                ["revision.author"] = "Eve",
                ["revision.id"] = "202"
            }));

        Assert.Contains("not supported via set", ex.Message);
        Assert.Empty(handler.Query("revision"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetRunMoveRevisions_QueryReturnsPairedMoveHalves()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var fromPath = handler.Add(paragraphPath, "run", null, new() { ["text"] = "old" });
        var toPath = handler.Add(paragraphPath, "run", null, new() { ["text"] = "new" });

        handler.Set(fromPath, new()
        {
            ["revision.type"] = "moveFrom",
            ["revision.author"] = "Mover",
            ["revision.id"] = "301"
        });
        handler.Set(toPath, new()
        {
            ["revision.type"] = "moveTo",
            ["revision.author"] = "Mover",
            ["revision.id"] = "301"
        });
        var moveFrom = Assert.Single(handler.Query("revision[@type=moveFrom]"));
        var moveTo = Assert.Single(handler.Query("revision[@type=moveTo]"));

        Assert.Equal("/revision[@id=301][@type=moveFrom]", moveFrom.Path);
        Assert.Equal("/revision[@id=301][@type=moveTo]", moveTo.Path);
        Assert.Equal("301", Fmt(moveFrom)["revision.id"]);
        Assert.Equal("301", Fmt(moveTo)["revision.id"]);
        Assert.Equal("old", moveFrom.Text);
        Assert.Equal("new", moveTo.Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetTableRowRevisionInsertion_QueryReturnsRowMarker()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        handler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["revision.type"] = "ins",
            ["revision.author"] = "Row Author",
            ["revision.id"] = "401"
        });
        var revision = Assert.Single(handler.Query("revision[@id=401]"));
        var row = handler.Get("/body/tbl[1]/tr[1]");

        Assert.Equal("rowIns", Fmt(revision)["revision.type"]);
        Assert.Equal("Row Author", Fmt(revision)["revision.author"]);
        Assert.Equal("401", Fmt(revision)["revision.id"]);
        Assert.Equal("ins", Fmt(row)["revision.type"]);
        Assert.Equal("Row Author", Fmt(row)["revision.author"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetTableAndCellRevisionFormat_QueryReturnsStructuralMarkers()
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
    public void RevisionActionAcceptAndRejectInsertion_UpdateDocumentText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var acceptedParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var acceptedRun = handler.Add(acceptedParagraph, "run", null, new() { ["text"] = "keep" });
        var rejectedParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var rejectedRun = handler.Add(rejectedParagraph, "run", null, new() { ["text"] = "drop" });

        handler.Set(acceptedRun, new()
        {
            ["revision.type"] = "ins",
            ["revision.author"] = "Reviewer",
            ["revision.id"] = "501"
        });
        handler.Set(rejectedRun, new()
        {
            ["revision.type"] = "ins",
            ["revision.author"] = "Reviewer",
            ["revision.id"] = "502"
        });

        handler.Set("/revision[@id=501]", new() { ["revision.action"] = "accept" });
        handler.Set("/revision[@id=502]", new() { ["revision.action"] = "reject" });

        Assert.Equal("keep", handler.Get(acceptedParagraph).Text);
        Assert.Equal("", handler.Get(rejectedParagraph).Text);
        Assert.Empty(handler.Query("revision"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void RevisionActionAcceptAndRejectDeletionAndFormat_UpdateCurrentDocument()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var acceptDelParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var acceptDelRun = handler.Add(acceptDelParagraph, "run", null, new() { ["text"] = "remove deleted" });
        var rejectDelParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var rejectDelRun = handler.Add(rejectDelParagraph, "run", null, new() { ["text"] = "restore deleted" });
        var acceptFormatParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var acceptFormatRun = handler.Add(acceptFormatParagraph, "run", null, new() { ["text"] = "keep bold" });
        var rejectFormatParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var rejectFormatRun = handler.Add(rejectFormatParagraph, "run", null, new() { ["text"] = "drop bold" });

        handler.Set(acceptDelRun, new() { ["revision.type"] = "del", ["revision.id"] = "701" });
        handler.Set(rejectDelRun, new() { ["revision.type"] = "del", ["revision.id"] = "702" });
        handler.Set(acceptFormatRun, new()
        {
            ["revision.type"] = "format",
            ["revision.id"] = "703",
            ["bold"] = "true"
        });
        handler.Set(rejectFormatRun, new()
        {
            ["revision.type"] = "format",
            ["revision.id"] = "704",
            ["bold"] = "true"
        });

        handler.Set("/revision[@id=701]", new() { ["revision.action"] = "accept" });
        handler.Set("/revision[@id=702]", new() { ["revision.action"] = "reject" });
        handler.Set("/revision[@id=703]", new() { ["revision.action"] = "accept" });
        handler.Set("/revision[@id=704]", new() { ["revision.action"] = "reject" });

        Assert.Equal("", handler.Get(acceptDelParagraph).Text);
        Assert.Equal("restore deleted", handler.Get(rejectDelParagraph).Text);
        Assert.Equal(true, Fmt(handler.Get(acceptFormatRun))["bold"]);
        Assert.False(Fmt(handler.Get(rejectFormatRun)).ContainsKey("bold"));
        Assert.Empty(handler.Query("revision"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void RevisionActionAcceptAndRejectMovePairs_UpdateDocumentText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var acceptedParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var acceptedFrom = handler.Add(acceptedParagraph, "run", null, new() { ["text"] = "old accepted" });
        var acceptedTo = handler.Add(acceptedParagraph, "run", null, new() { ["text"] = "new accepted" });
        var rejectedParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var rejectedFrom = handler.Add(rejectedParagraph, "run", null, new() { ["text"] = "old rejected" });
        var rejectedTo = handler.Add(rejectedParagraph, "run", null, new() { ["text"] = "new rejected" });

        handler.Set(acceptedFrom, new() { ["revision.type"] = "moveFrom", ["revision.id"] = "801" });
        handler.Set(acceptedTo, new() { ["revision.type"] = "moveTo", ["revision.id"] = "801" });
        handler.Set(rejectedFrom, new() { ["revision.type"] = "moveFrom", ["revision.id"] = "802" });
        handler.Set(rejectedTo, new() { ["revision.type"] = "moveTo", ["revision.id"] = "802" });

        handler.Set("/revision[@id=801][@type=moveFrom]", new() { ["revision.action"] = "accept" });
        handler.Set("/revision[@id=801][@type=moveTo]", new() { ["revision.action"] = "accept" });
        handler.Set("/revision[@id=802][@type=moveFrom]", new() { ["revision.action"] = "reject" });
        handler.Set("/revision[@id=802][@type=moveTo]", new() { ["revision.action"] = "reject" });

        Assert.Equal("new accepted", handler.Get(acceptedParagraph).Text);
        Assert.Equal("old rejected", handler.Get(rejectedParagraph).Text);
        Assert.Empty(handler.Query("revision"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetSectionRevisionFormat_ReadsBackSectPrChange()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);

        handler.Set("/section[1]", new()
        {
            ["revision.type"] = "format",
            ["revision.author"] = "Section Author",
            ["revision.id"] = "601",
            ["orientation"] = "landscape"
        });
        var revision = Assert.Single(handler.Query("revision[@id=601]"));
        var section = handler.Get("/section[1]");

        Assert.Equal("format", Fmt(revision)["revision.type"]);
        Assert.Equal("Section Author", Fmt(revision)["revision.author"]);
        Assert.Equal("601", Fmt(revision)["revision.id"]);
        Assert.Equal("landscape", Fmt(section)["orientation"]);
        Assert.Equal("Section Author", Fmt(section)["sectPrChange.author"]);
        Assert.Equal("601", Fmt(section)["sectPrChange.id"]);
        Assert.Empty(handler.Validate());
    }
}
