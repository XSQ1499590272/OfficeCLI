using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordAliasContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void RepresentativeLegacyAliases_WriteAndReadCanonicalKeys()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/", new() { ["creator"] = "Alias author" });
        Assert.Equal("Alias author", Fmt(handler.Get("/"))["author"]);

        var paragraphPath = handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Alias paragraph",
            ["alignment"] = "right"
        });
        Assert.Equal("right", Fmt(handler.Get(paragraphPath))["align"]);

        var runPath = handler.Add(paragraphPath, "run", null, new()
        {
            ["text"] = "Alias run",
            ["strikethrough"] = "true"
        });
        Assert.Equal(true, Fmt(handler.Get(runPath))["strike"]);

        var hyperlinkPath = handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["text"] = "Alias link",
            ["href"] = "https://example.com/alias"
        });
        Assert.Equal("https://example.com/alias", Fmt(handler.Get(hyperlinkPath))["url"]);

        handler.Add("/body", "table", null, new() { ["data"] = "A,B,C" });
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["gridspan"] = "2" });
        var mergedCell = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        Assert.Equal(2, Fmt(mergedCell)["colspan"]);

        handler.Add(paragraphPath, "field", null, new()
        {
            ["type"] = "page",
            ["text"] = "1"
        });
        var field = Assert.Single(handler.Query("field"));
        Assert.Equal("page", Fmt(field)["fieldType"]);
        Assert.Empty(handler.Validate());
    }
}
