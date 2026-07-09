using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordTextboxContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddTextbox_ReadsBackAddressableContentTree()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var textboxPath = handler.Add("/body", "textbox", null, new()
        {
            ["text"] = "Sidebar note",
            ["width"] = "4cm",
            ["height"] = "2cm",
            ["geometry"] = "roundRect",
            ["fill"] = "EAF2FF",
            ["line.color"] = "2B579A",
            ["line.width"] = "1pt",
            ["hPosition"] = "2cm",
            ["vPosition"] = "1cm",
            ["hRelative"] = "page",
            ["vRelative"] = "paragraph",
            ["rotation"] = "15",
            ["textDirection"] = "vert",
            ["name"] = "Callout"
        });

        var textbox = handler.Get(textboxPath, depth: 1);
        var paragraph = handler.Get($"{textboxPath}/p[1]");
        var run = handler.Get($"{textboxPath}/p[1]/r[1]");

        Assert.Equal("/body/textbox[1]", textboxPath);
        Assert.Equal("txbxContent", textbox.Type);
        Assert.Equal("Sidebar note", textbox.Text);
        Assert.Equal(1, textbox.ChildCount);
        Assert.Single(textbox.Children);
        Assert.Equal($"{textboxPath}/p[1]", textbox.Children[0].Path);
        Assert.Equal("paragraph", paragraph.Type);
        Assert.Equal("Sidebar note", paragraph.Text);
        Assert.Equal("run", run.Type);
        Assert.Equal("Sidebar note", run.Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetTextbox_SupportsShapeSurfaceButKeepsTextAndPositionAddOnly()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var textboxPath = handler.Add("/body", "textbox", null, new()
        {
            ["text"] = "Original text",
            ["width"] = "4cm",
            ["height"] = "2cm"
        });

        var unsupported = handler.Set(textboxPath, new()
        {
            ["fill"] = "FFECEC",
            ["line.color"] = "C00000",
            ["line.width"] = "2pt",
            ["width"] = "5cm",
            ["height"] = "3cm",
            ["geometry"] = "ellipse"
        });
        var textUnsupported = handler.Set(textboxPath, new() { ["text"] = "Changed text" });
        var positionError = Assert.Throws<ArgumentException>(() =>
            handler.Set(textboxPath, new() { ["hPosition"] = "3cm" }));
        var paragraph = handler.Get($"{textboxPath}/p[1]");

        Assert.Empty(unsupported);
        Assert.Contains("text", textUnsupported);
        Assert.Contains("position keys", positionError.Message);
        Assert.Equal("Original text", paragraph.Text);
        Assert.Empty(handler.Validate());
    }
}
