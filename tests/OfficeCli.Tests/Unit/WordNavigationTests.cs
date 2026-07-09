using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordNavigationTests : WordTestBase
{
    [Fact]
    public void PositionalPaths_NavigateBodyRunTableRowAndCell()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paragraphPath, "run", null, new() { ["text"] = "path run" });
        handler.Add("/body", "table", null, new() { ["data"] = "A,B;C,D" });

        Assert.Equal("paragraph", handler.Get("/body/p[1]").Type);
        Assert.Equal("run", handler.Get("/body/p[1]/r[1]").Type);
        Assert.Equal("path run", handler.Get(runPath).Text);
        Assert.Equal("table", handler.Get("/body/tbl[1]").Type);
        Assert.Equal("row", handler.Get("/body/tbl[1]/tr[2]").Type);
        Assert.Equal("D", handler.Get("/body/tbl[1]/tr[2]/tc[2]").Text);
    }

    [Fact]
    public void StableParagraphPath_NavigatesAfterReopen()
    {
        var path = CreateBlankDocx();
        string paragraphPath;
        string runPath;

        using (var handler = new WordHandler(path, editable: true))
        {
            paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
            runPath = handler.Add(paragraphPath, "run", null, new() { ["text"] = "stable text" });
            handler.Add(paragraphPath, "bookmark", null, new() { ["name"] = "StableMark" });
        }

        using var reopened = new WordHandler(path, editable: false);
        var paragraph = reopened.Get(paragraphPath);
        var run = reopened.Get(runPath);
        var bookmark = Assert.Single(reopened.Query("bookmark"));

        Assert.Contains("@paraId=", paragraphPath);
        Assert.Equal("paragraph", paragraph.Type);
        Assert.Equal("run", run.Type);
        Assert.Equal("stable text", run.Text);
        Assert.Equal("StableMark", Fmt(bookmark)["name"]);
    }
}
