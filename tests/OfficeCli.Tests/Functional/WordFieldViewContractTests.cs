using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordFieldViewContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void ViewTextAndIssues_DistinguishUnevaluatedAndDirtyFieldCaches()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "field: " });
        handler.Add(paragraphPath, "field", null, new() { ["fieldType"] = "page", ["text"] = "" });
        var dirtyParagraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "dirty field: " });
        handler.Add(dirtyParagraphPath, "field", null, new() { ["fieldType"] = "page", ["text"] = "1" });

        var unevaluated = handler.ViewAsIssues("field_not_evaluated");
        Assert.NotEmpty(unevaluated);
        var unevaluatedIssue = Assert.Single(unevaluated);
        Assert.Equal("field_not_evaluated", unevaluatedIssue.Subtype);
        Assert.Contains("PAGE", unevaluatedIssue.Context);
        Assert.Contains("#OCLI_NOTEVAL!{PAGE}", handler.ViewAsText());

        var instructionPath = handler.Query("instrText")[1].Path;
        handler.Set(instructionPath, new() { ["instr"] = "NUMPAGES" });

        var stale = handler.ViewAsIssues("field_cache_stale");
        var staleIssue = Assert.Single(stale);
        Assert.Equal("field_cache_stale", staleIssue.Subtype);
        Assert.Contains("NUMPAGES", staleIssue.Context);
        Assert.Contains("#OCLI_NOTEVAL!{NUMPAGES}", handler.ViewAsText());
        Assert.Empty(handler.Validate());
    }
}
