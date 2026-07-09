using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordPermissionContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddPermStartAndEnd_ReadsBackPairAttributes()
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
        var paragraph = handler.Get(paragraphPath, depth: 1);
        var startFmt = Fmt(start);
        var endFmt = Fmt(end);

        Assert.EndsWith("/permStart[@id=7]", startPath);
        Assert.EndsWith("/permEnd[@id=7]", endPath);
        Assert.Equal("permStart", start.Type);
        Assert.Equal("7", startFmt["id"]);
        Assert.Equal("everyone", startFmt["edGrp"]);
        Assert.Equal("jane@example.com", startFmt["ed"]);
        Assert.Equal("0", startFmt["colFirst"]);
        Assert.Equal("2", startFmt["colLast"]);
        Assert.Equal("permEnd", end.Type);
        Assert.Equal("7", endFmt["id"]);
        Assert.Contains(paragraph.Children, child => child.Type == "permStart");
        Assert.Contains(paragraph.Children, child => child.Type == "permEnd");
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void RemovePermMarkers_DeletesOnlyMarkers()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "editable text" });
        var startPath = handler.Add(paragraphPath, "permStart", null, new() { ["id"] = "7" });
        var endPath = handler.Add(paragraphPath, "permEnd", null, new() { ["id"] = "7" });

        handler.Remove($"{paragraphPath}/permEnd[1]");
        handler.Remove($"{paragraphPath}/permStart[1]");
        var paragraph = handler.Get(paragraphPath, depth: 1);

        Assert.DoesNotContain(paragraph.Children, child => child.Type == "permStart");
        Assert.DoesNotContain(paragraph.Children, child => child.Type == "permEnd");
        Assert.Equal("editable text", paragraph.Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddPermStart_WithoutIntegerIdThrowsBeforeMutation()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "editable text" });

        var missing = Assert.Throws<ArgumentException>(() =>
            handler.Add(paragraphPath, "permStart", null, new()));
        var nonInteger = Assert.Throws<ArgumentException>(() =>
            handler.Add(paragraphPath, "permStart", null, new() { ["id"] = "abc" }));
        var paragraph = handler.Get(paragraphPath, depth: 1);

        Assert.Contains("id", missing.Message);
        Assert.Contains("id", nonInteger.Message);
        Assert.DoesNotContain(paragraph.Children, child => child.Type == "permStart");
        Assert.Empty(handler.Validate());
    }
}
