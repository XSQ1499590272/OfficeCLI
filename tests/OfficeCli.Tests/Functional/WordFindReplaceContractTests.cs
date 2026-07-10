using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordFindReplaceContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void SetFindReplace_ReplacesOnlyMatchingParagraphs()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.Add("/body", "paragraph", null, new() { ["text"] = "needle here" });
        var second = handler.Add("/body", "paragraph", null, new() { ["text"] = "plain here" });

        var unsupported = handler.Set("/body", new()
        {
            ["find"] = "needle",
            ["replace"] = "thread"
        });

        Assert.Empty(unsupported);
        Assert.Equal(1, handler.LastFindMatchCount);
        Assert.Equal("thread here", handler.Get(first).Text);
        Assert.Equal("plain here", handler.Get(second).Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetRegexFindReplace_ExpandsRegexBackreferences()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Invoice 42" });

        handler.Set(paragraphPath, new()
        {
            ["find"] = "(\\d+)",
            ["regex"] = "true",
            ["replace"] = "#$1"
        });

        Assert.Equal("Invoice #42", handler.Get(paragraphPath).Text);
        Assert.Equal(1, handler.LastFindMatchCount);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetFindReplaceWithRevisionAuthor_CreatesDeletionAndInsertionPerMatch()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "old old" });

        handler.Set(paragraphPath, new()
        {
            ["find"] = "old",
            ["replace"] = "new",
            ["revision.author"] = "Tracked Replace",
            ["revision.date"] = "2026-01-02T03:04:05Z"
        });

        Assert.Equal(2, handler.LastFindMatchCount);
        var deleted = handler.Query("revision[@author='Tracked Replace'][@type=del]");
        var inserted = handler.Query("revision[@author='Tracked Replace'][@type=ins]");
        Assert.Equal(2, deleted.Count);
        Assert.Equal(2, inserted.Count);
        Assert.All(deleted, revision => Assert.Equal("old", revision.Text));
        Assert.All(inserted, revision => Assert.Equal("new", revision.Text));
        Assert.All(deleted.Concat(inserted), revision =>
        {
            Assert.Equal("Tracked Replace", Fmt(revision)["revision.author"]);
            Assert.Equal(
                DateTimeOffset.Parse("2026-01-02T03:04:05Z"),
                DateTimeOffset.Parse(Fmt(revision)["revision.date"]!.ToString()!).ToUniversalTime());
            Assert.False(string.IsNullOrWhiteSpace(Fmt(revision)["revision.id"]?.ToString()));
        });
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetFindWithRevisionAuthor_EmptyReplacementCreatesDeletionOnly()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "remove this" });

        handler.Set(paragraphPath, new()
        {
            ["find"] = "remove",
            ["replace"] = "",
            ["revision.author"] = "Tracked Delete"
        });

        var deletion = Assert.Single(handler.Query("revision[@author='Tracked Delete'][@type=del]"), _ => true);
        Assert.Equal("remove", deletion.Text);
        Assert.Empty(handler.Query("revision[@author='Tracked Delete'][@type=ins]"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetFindWithRevisionAuthor_FormatOnlyCreatesFormatMarker()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "format this" });

        handler.Set(paragraphPath, new()
        {
            ["find"] = "format",
            ["bold"] = "true",
            ["revision.author"] = "Tracked Format"
        });

        var revision = Assert.Single(handler.Query("revision[@author='Tracked Format'][@type=format]"));
        var run = Assert.Single(handler.Get(paragraphPath, depth: 1).Children,
            child => child.Type == "run" && child.Text == "format");
        Assert.Equal("Tracked Format", Fmt(revision)["revision.author"]);
        Assert.Equal(true, Fmt(run)["bold"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedSet_BareFindWithoutActionThrows()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "needle" });

        var ex = Assert.Throws<ArgumentException>(() =>
            handler.Set("/body", new() { ["find"] = "needle" }));

        Assert.Contains("'find' requires", ex.Message);
    }

    [Fact]
    public void SetRange_FormatsHalfOpenSpanAcrossBodyParagraphs()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.Add("/body", "paragraph", null, new() { ["text"] = "abcdef" });
        var second = handler.Add("/body", "paragraph", null, new() { ["text"] = "ghijkl" });

        var unsupported = handler.Set("/body", new()
        {
            ["range"] = "2:8",
            ["bold"] = "true",
            ["color"] = "00AA00"
        });

        Assert.Empty(unsupported);
        Assert.Equal(2, handler.LastFindMatchCount);

        var firstRuns = handler.Get(first, depth: 1).Children.Where(child => child.Type == "run").ToList();
        var secondRuns = handler.Get(second, depth: 1).Children.Where(child => child.Type == "run").ToList();

        Assert.Equal(["ab", "cdef"], firstRuns.Select(run => run.Text ?? "").ToArray());
        Assert.Equal(true, Fmt(firstRuns[1])["bold"]);
        Assert.Equal("#00AA00", Fmt(firstRuns[1])["color"]);
        Assert.Equal(["gh", "ijkl"], secondRuns.Select(run => run.Text ?? "").ToArray());
        Assert.Equal(true, Fmt(secondRuns[0])["bold"]);
        Assert.Equal("#00AA00", Fmt(secondRuns[0])["color"]);
        Assert.False(Fmt(secondRuns[1]).ContainsKey("bold"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedSet_RangeRejectsFindTextAndMissingFormatProps()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "abcdef" });

        var findAndRange = Assert.Throws<ArgumentException>(() =>
            handler.Set(paragraphPath, new() { ["find"] = "bc", ["range"] = "1:2", ["bold"] = "true" }));
        Assert.Contains("mutually exclusive", findAndRange.Message);

        var textRange = Assert.Throws<ArgumentException>(() =>
            handler.Set(paragraphPath, new() { ["range"] = "1:2", ["text"] = "x" }));
        Assert.Contains("formatting only", textRange.Message);

        var missingFormat = Assert.Throws<ArgumentException>(() =>
            handler.Set(paragraphPath, new() { ["range"] = "1:2" }));
        Assert.Contains("requires format properties", missingFormat.Message);
    }
}
