using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordDiagramContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddDiagramNative_ReturnsGroupAndExposesNodeTextboxes()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var groupPath = handler.Add("/body", "diagram", null, new()
        {
            ["render"] = "native",
            ["mermaid"] = "flowchart TD; A[Start] --> B[Done]",
            ["width"] = "8cm"
        });

        var group = handler.Get(groupPath);
        var firstNode = handler.Get("/body/textbox[1]/p[1]");
        var fmt = Fmt(group);

        Assert.Equal("/body/group[1]", groupPath);
        Assert.Equal("group", group.Type);
        Assert.Equal("/body/group[1]", group.Path);
        Assert.True(fmt.ContainsKey("x"));
        Assert.True(fmt.ContainsKey("y"));
        Assert.True(fmt.ContainsKey("width"));
        Assert.True(fmt.ContainsKey("height"));
        Assert.Equal("Start", firstNode.Text);
        Assert.Throws<ArgumentException>(() => handler.Get("/body/textbox[2]/p[1]"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetAndRemoveDiagramGroup_UpdatesGroupSurfaceAndDeletesWrapper()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var groupPath = handler.Add("/body", "diagram", null, new()
        {
            ["render"] = "native",
            ["mermaid"] = "flowchart TD; A[Start] --> B[Done]"
        });

        var unsupported = handler.Set(groupPath, new()
        {
            ["width"] = "6cm",
            ["height"] = "3cm"
        });
        var resized = handler.Get(groupPath);

        Assert.Empty(unsupported);
        Assert.Equal("6cm", Fmt(resized)["width"]);
        Assert.Equal("3cm", Fmt(resized)["height"]);

        handler.Remove(groupPath);

        Assert.Throws<ArgumentException>(() => handler.Get(groupPath));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddDiagram_WithoutSourceThrowsBeforeMutation()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var ex = Assert.Throws<ArgumentException>(() =>
            handler.Add("/body", "diagram", null, new() { ["render"] = "native" }));

        Assert.Contains("diagram requires", ex.Message);
        Assert.Throws<ArgumentException>(() => handler.Get("/body/group[1]"));
    }
}
