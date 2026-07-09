using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordFieldContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddField_ReadsCollapsedFieldInstrTextAndFieldChars()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(paragraphPath, "field", null, new()
        {
            ["fieldType"] = "page",
            ["text"] = "7",
            ["fldLock"] = "true"
        });

        var field = Assert.Single(handler.Query("field"));
        var instrText = Assert.Single(handler.Query("instrText"));
        var fieldChars = handler.Query("fieldChar");

        Assert.Equal("field", field.Type);
        Assert.Equal("7", field.Text);
        Assert.Equal("PAGE", Fmt(field)["instruction"]);
        Assert.Equal("page", Fmt(field)["fieldType"]);
        Assert.Equal(true, Fmt(field)["evaluated"]);
        Assert.Equal(true, Fmt(field)["fldLock"]);
        Assert.Contains("PAGE", instrText.Text);
        Assert.Contains(fieldChars, node => Equals(Fmt(node)["fieldCharType"], "begin"));
        Assert.Contains(fieldChars, node => Equals(Fmt(node)["fieldCharType"], "separate"));
        Assert.Contains(fieldChars, node => Equals(Fmt(node)["fieldCharType"], "end"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetInstrText_UpdatesOwningFieldInstructionAndMarksDirty()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(paragraphPath, "field", null, new()
        {
            ["fieldType"] = "page",
            ["text"] = "1"
        });
        var instrPath = Assert.Single(handler.Query("instrText")).Path;

        handler.Set(instrPath, new() { ["instr"] = "NUMPAGES" });
        var field = handler.Get("/field[1]");

        Assert.Equal("NUMPAGES", Fmt(field)["instruction"]);
        Assert.Equal("numpages", Fmt(field)["fieldType"]);
        Assert.Equal(true, Fmt(field)["dirty"]);
    }

    [Fact]
    public void AddToc_ReadsBackLevelsSwitchesAndTitle()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var tocPath = handler.Add("/body", "toc", null, new()
        {
            ["levels"] = "1-2",
            ["title"] = "Contents",
            ["hyperlinks"] = "false",
            ["pageNumbers"] = "false"
        });

        var toc = handler.Get(tocPath);
        var queried = Assert.Single(handler.Query("toc"));

        Assert.Equal("toc", toc.Type);
        Assert.Contains("TOC", toc.Text);
        Assert.Equal("1-2", Fmt(toc)["levels"]);
        Assert.Equal("Contents", Fmt(toc)["title"]);
        Assert.Equal(false, Fmt(toc)["hyperlinks"]);
        Assert.Equal(false, Fmt(toc)["pageNumbers"]);
        Assert.Equal("/toc[1]", queried.Path);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddSeqFields_CachesBodyOrderNumbers()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var firstParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var secondParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });

        handler.Add(firstParagraph, "field", null, new()
        {
            ["fieldType"] = "seq",
            ["identifier"] = "Figure"
        });
        handler.Add(secondParagraph, "field", null, new()
        {
            ["fieldType"] = "seq",
            ["identifier"] = "Figure"
        });
        var fields = handler.Query("field");

        Assert.Equal(2, fields.Count);
        Assert.Equal("1", fields[0].Text);
        Assert.Equal("2", fields[1].Text);
        Assert.Equal("SEQ Figure", Fmt(fields[0])["instruction"]?.ToString()?.Trim());
        Assert.Equal("seq", Fmt(fields[0])["fieldType"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddSeqFields_WithResetAndRomanSwitchesCachesCurrentEvaluatorResults()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var firstParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var secondParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });

        handler.Add(firstParagraph, "field", null, new()
        {
            ["fieldType"] = "seq",
            ["identifier"] = "Figure",
            ["switches"] = @"\r 5 \* ROMAN"
        });
        handler.Add(secondParagraph, "field", null, new()
        {
            ["fieldType"] = "seq",
            ["identifier"] = "Figure",
            ["switches"] = @"\* ROMAN"
        });
        var fields = handler.Query("field");

        Assert.Equal(2, fields.Count);
        Assert.Equal("V", fields[0].Text);
        Assert.Equal("II", fields[1].Text);
        Assert.Equal(@"SEQ Figure \r 5 \* ROMAN", Fmt(fields[0])["instruction"]?.ToString()?.Trim());
        Assert.Equal(@"SEQ Figure \* ROMAN", Fmt(fields[1])["instruction"]?.ToString()?.Trim());
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetRootRecalcFieldsSeq_RewritesSeqFieldCachesFromCurrentInstructions()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var firstParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var secondParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });

        handler.Add(firstParagraph, "field", null, new()
        {
            ["fieldType"] = "seq",
            ["identifier"] = "Figure"
        });
        handler.Add(secondParagraph, "field", null, new()
        {
            ["fieldType"] = "seq",
            ["identifier"] = "Figure"
        });
        var secondInstrPath = handler.Query("instrText")[1].Path;

        handler.Set(secondInstrPath, new() { ["instr"] = "SEQ Table" });
        handler.Set("/", new() { ["recalcFields"] = "seq" });
        var fields = handler.Query("field");

        Assert.Equal("1", fields[0].Text);
        Assert.Equal("1", fields[1].Text);
        Assert.Equal("SEQ Table", Fmt(fields[1])["instruction"]?.ToString()?.Trim());
        Assert.Empty(handler.Validate());
    }
}
