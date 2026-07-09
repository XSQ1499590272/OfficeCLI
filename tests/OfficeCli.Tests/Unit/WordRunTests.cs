using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordRunTests : WordTestBase
{
    [Fact]
    public void AddRun_ReadsBackTextAndCanonicalRunFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paraPath, "run", null, new()
        {
            ["text"] = "Formatted run",
            ["bold"] = "true",
            ["italic"] = "true",
            ["font"] = "Georgia",
            ["size"] = "14pt",
            ["color"] = "C00000",
            ["underline"] = "single"
        });

        var node = handler.Get(runPath);

        Assert.Equal("run", node.Type);
        Assert.Equal("Formatted run", node.Text);
        Assert.Equal(true, Fmt(node)["bold"]);
        Assert.Equal(true, Fmt(node)["italic"]);
        Assert.Equal("Georgia", Fmt(node)["font.latin"]);
        Assert.Equal("14pt", Fmt(node)["size"]);
        Assert.Equal("#C00000", Fmt(node)["color"]);
        Assert.Equal("single", Fmt(node)["underline"]);
    }

    [Fact]
    public void SetRunText_PreservesExistingRunFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paraPath, "run", null, new()
        {
            ["text"] = "before",
            ["bold"] = "true",
            ["color"] = "2A9D8F"
        });

        handler.Set(runPath, new() { ["text"] = "after" });

        var node = handler.Get(runPath);
        Assert.Equal("after", node.Text);
        Assert.Equal(true, Fmt(node)["bold"]);
        Assert.Equal("#2A9D8F", Fmt(node)["color"]);
    }

    [Fact]
    public void SetRunText_ToEmptyKeepsRunAndFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paraPath, "run", null, new()
        {
            ["text"] = "clear me",
            ["italic"] = "true",
            ["underline"] = "single"
        });

        handler.Set(runPath, new() { ["text"] = "" });

        var node = handler.Get(runPath);
        Assert.Equal("run", node.Type);
        Assert.Equal("", node.Text);
        Assert.Equal(true, Fmt(node)["italic"]);
        Assert.Equal("single", Fmt(node)["underline"]);
    }

    [Fact]
    public void AddRun_ReadsBackScriptSpecificFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paraPath, "run", null, new()
        {
            ["text"] = "Mixed ابج 文",
            ["font.latin"] = "Aptos",
            ["font.ea"] = "Songti SC",
            ["font.cs"] = "Arial",
            ["lang.latin"] = "en-US",
            ["lang.ea"] = "zh-CN",
            ["lang.cs"] = "ar-SA",
            ["size.cs"] = "16pt",
            ["bold.cs"] = "true",
            ["italic.cs"] = "true",
            ["rtl"] = "true"
        });

        var node = handler.Get(runPath);

        Assert.Equal("Aptos", Fmt(node)["font.latin"]);
        Assert.Equal("Songti SC", Fmt(node)["font.ea"]);
        Assert.Equal("Arial", Fmt(node)["font.cs"]);
        Assert.Equal("en-US", Fmt(node)["lang.latin"]);
        Assert.Equal("zh-CN", Fmt(node)["lang.ea"]);
        Assert.Equal("ar-SA", Fmt(node)["lang.cs"]);
        Assert.Equal("16pt", Fmt(node)["size.cs"]);
        Assert.Equal(true, Fmt(node)["bold.cs"]);
        Assert.Equal(true, Fmt(node)["italic.cs"]);
        Assert.Equal("rtl", Fmt(node)["direction"]);
    }
}
