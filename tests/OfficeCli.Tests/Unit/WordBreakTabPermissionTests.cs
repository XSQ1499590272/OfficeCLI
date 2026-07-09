using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordBreakTabPermissionTests : WordTestBase
{
    [Fact]
    public void AddPageBreak_ReadsBackBreakTypeAndQueryPath()
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

    [Fact]
    public void TabStop_AddSetRemoveAndPositionalTab_ReadBackCurrentBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "tab host" });
        var tabPath = handler.Add(paragraphPath, "tab", null, new()
        {
            ["pos"] = "1440",
            ["val"] = "right",
            ["leader"] = "dot"
        });
        var ptabPath = handler.Add(paragraphPath, "ptab", null, new()
        {
            ["align"] = "right",
            ["relativeTo"] = "margin",
            ["leader"] = "dot"
        });

        handler.Set(tabPath, new()
        {
            ["pos"] = "720",
            ["val"] = "center",
            ["leader"] = "hyphen"
        });
        handler.Set(ptabPath, new()
        {
            ["align"] = "center",
            ["leader"] = "underscore"
        });

        var tab = Assert.Single((IEnumerable<Dictionary<string, object?>>)Fmt(handler.Get(paragraphPath))["tabs"]!);
        var ptab = handler.Get(ptabPath);

        Assert.Equal(720, Convert.ToInt32(tab["pos"]));
        Assert.Equal("center", tab["val"]);
        Assert.Equal("hyphen", tab["leader"]);
        Assert.Equal("center", Fmt(ptab)["align"]);
        Assert.Equal("underscore", Fmt(ptab)["leader"]);

        handler.Remove(tabPath);
        Assert.False(Fmt(handler.Get(paragraphPath)).ContainsKey("tabs"));
        Assert.Contains(handler.Validate(), error =>
            error.ErrorType == "Schema" && error.Description.Contains("incomplete content"));
    }

    [Fact]
    public void PermissionMarkers_ReadBackRemoveAndRejectInvalidIds()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "editable text" });
        var startPath = handler.Add(paragraphPath, "permStart", null, new()
        {
            ["id"] = "7",
            ["edGrp"] = "everyone",
            ["ed"] = "jane@example.com",
            ["colFirst"] = "0",
            ["colLast"] = "2"
        });
        var endPath = handler.Add(paragraphPath, "permEnd", null, new() { ["id"] = "7" });

        var start = handler.Get($"{paragraphPath}/permStart[1]");
        var end = handler.Get($"{paragraphPath}/permEnd[1]");

        Assert.EndsWith("/permStart[@id=7]", startPath);
        Assert.EndsWith("/permEnd[@id=7]", endPath);
        Assert.Equal("7", Fmt(start)["id"]);
        Assert.Equal("everyone", Fmt(start)["edGrp"]);
        Assert.Equal("jane@example.com", Fmt(start)["ed"]);
        Assert.Equal("0", Fmt(start)["colFirst"]);
        Assert.Equal("2", Fmt(start)["colLast"]);
        Assert.Equal("7", Fmt(end)["id"]);

        handler.Remove($"{paragraphPath}/permEnd[1]");
        handler.Remove($"{paragraphPath}/permStart[1]");
        var paragraph = handler.Get(paragraphPath, depth: 1);

        Assert.DoesNotContain(paragraph.Children, child => child.Type == "permStart");
        Assert.DoesNotContain(paragraph.Children, child => child.Type == "permEnd");
        Assert.Equal("editable text", paragraph.Text);
        Assert.Throws<ArgumentException>(() => handler.Add(paragraphPath, "permStart", null, new()));
        Assert.Throws<ArgumentException>(() => handler.Add(paragraphPath, "permStart", null, new() { ["id"] = "abc" }));
        Assert.Empty(handler.Validate());
    }
}
