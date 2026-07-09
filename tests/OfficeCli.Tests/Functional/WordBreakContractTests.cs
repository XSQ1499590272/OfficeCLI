using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordBreakContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddPageBreak_AcceptsCanonicalBreakTypeAndQueriesAsBreak()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var breakPath = handler.Add(paragraphPath, "pagebreak", null, new() { ["breakType"] = "column" });

        var breakNode = handler.Get(breakPath);
        var queried = Assert.Single(handler.Query("pagebreak"));

        Assert.Equal("break", breakNode.Type);
        Assert.Equal("column", Fmt(breakNode)["breakType"]);
        Assert.Equal(breakPath, queried.Path);
        Assert.Equal("column", Fmt(queried)["breakType"]);
    }
}
