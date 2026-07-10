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

    [Fact]
    public void AddRun_ReadsBackAdvancedToggleFormatting_AndExplicitFalseOverrides()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paraPath, "run", null, new()
        {
            ["text"] = "Advanced",
            ["highlight"] = "yellow",
            ["strike"] = "true",
            ["dstrike"] = "true",
            ["caps"] = "true",
            ["smallcaps"] = "true",
            ["vanish"] = "true",
            ["outline"] = "true",
            ["shadow"] = "true",
            ["emboss"] = "true",
            ["imprint"] = "true",
            ["noproof"] = "true"
        });

        var format = Fmt(handler.Get(runPath));
        Assert.Equal("yellow", format["highlight"]);
        foreach (var key in new[] { "strike", "dstrike", "caps", "smallcaps", "vanish", "outline", "shadow", "emboss", "imprint", "noproof" })
            Assert.Equal(true, format[key]);

        handler.Set(runPath, new()
        {
            ["strike"] = "false",
            ["dstrike"] = "false",
            ["caps"] = "false",
            ["smallcaps"] = "false",
            ["vanish"] = "false",
            ["outline"] = "false",
            ["shadow"] = "false",
            ["emboss"] = "false",
            ["imprint"] = "false",
            ["noproof"] = "false",
            ["highlight"] = "none"
        });

        format = Fmt(handler.Get(runPath));
        Assert.Equal("none", format["highlight"]);
        foreach (var key in new[] { "strike", "dstrike", "caps", "smallcaps", "vanish", "outline", "shadow", "emboss", "imprint", "noproof" })
            Assert.Equal(false, format[key]);
    }

    [Fact]
    public void AddRun_ReadsBackW14TextEffects_AndRejectsInvalidSpecs()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var runPath = handler.Add(paraPath, "run", null, new()
        {
            ["text"] = "Effects",
            ["textOutline"] = "1pt;FF0000",
            ["textFill"] = "00FF00",
            ["w14shadow"] = "0000FF",
            ["w14glow"] = "FF00FF;6;50",
            ["w14reflection"] = "full"
        });

        var format = Fmt(handler.Get(runPath));
        Assert.Equal("1pt;#FF0000", format["textOutline"]);
        Assert.Equal("#00FF00", format["textFill"]);
        Assert.Equal("#0000FF;4;45;3;40", format["w14shadow"]);
        Assert.Equal("#FF00FF;6;50", format["w14glow"]);
        Assert.Equal("full", format["w14reflection"]);
        Assert.Throws<ArgumentException>(() =>
            handler.Add(paraPath, "run", null, new() { ["text"] = "bad", ["textOutline"] = "wide;FF0000" }));
    }

    [Fact]
    public void AddRun_ReadsBackShadingAndThemeLinkedColor()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var solidPath = handler.Add(paraPath, "run", null, new()
        {
            ["text"] = "Solid",
            ["fill"] = "FFFF00",
            ["color"] = "FFFFFF;themeColor=background1;themeTint=99"
        });
        var patternPath = handler.Add(paraPath, "run", null, new()
        {
            ["text"] = "Pattern",
            ["shading"] = "pct20;FFFF00;0000FF;themeFill=accent1;themeColor=accent2"
        });

        var solid = Fmt(handler.Get(solidPath));
        Assert.Equal("#FFFF00", solid["fill"]);
        Assert.Equal("#FFFFFF;themeColor=background1;themeTint=99", solid["color"]);

        var pattern = Fmt(handler.Get(patternPath));
        Assert.Equal("pct20", pattern["shading.val"]);
        Assert.Equal("#FFFF00", pattern["shading.fill"]);
        Assert.Equal("#0000FF", pattern["shading.color"]);
        Assert.Equal("accent1", pattern["shading.themeFill"]);
        Assert.Equal("accent2", pattern["shading.themeColor"]);
    }

    [Fact]
    public void SetParagraphRange_FormatsOnlyHalfOpenCharacterSpan()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "abcdef" });

        handler.Set(paragraphPath, new()
        {
            ["range"] = "1:4",
            ["bold"] = "true",
            ["color"] = "FF0000"
        });

        var paragraph = handler.Get(paragraphPath, depth: 1);
        var runs = paragraph.Children.Where(child => child.Type == "run").ToList();
        Assert.Equal(new[] { "a", "bcd", "ef" }, runs.Select(r => r.Text).ToArray());
        Assert.False(Fmt(runs[0]).ContainsKey("bold"));
        Assert.Equal(true, Fmt(runs[1])["bold"]);
        Assert.Equal("#FF0000", Fmt(runs[1])["color"]);
        Assert.False(Fmt(runs[2]).ContainsKey("bold"));
        Assert.Equal(1, handler.LastFindMatchCount);

        Assert.Throws<ArgumentException>(() =>
            handler.Set(paragraphPath, new() { ["find"] = "bc", ["range"] = "1:2", ["bold"] = "true" }));
        Assert.Throws<ArgumentException>(() =>
            handler.Set(paragraphPath, new() { ["range"] = "99:100", ["bold"] = "true" }));
        Assert.Throws<ArgumentException>(() =>
            handler.Set(paragraphPath, new() { ["range"] = "1:2", ["text"] = "x" }));
    }

    [Fact]
    public void SetParagraphRange_FormatsMultipleDisjointSpansInPositionOrder()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "abcdefghij" });

        handler.Set(paragraphPath, new()
        {
            ["range"] = "7:9,1:3",
            ["bold"] = "true"
        });

        var runs = handler.Get(paragraphPath, depth: 1).Children
            .Where(child => child.Type == "run").ToList();
        Assert.Equal(["a", "bc", "defg", "hi", "j"], runs.Select(run => run.Text ?? "").ToArray());
        Assert.False(Fmt(runs[0]).ContainsKey("bold"));
        Assert.Equal(true, Fmt(runs[1])["bold"]);
        Assert.False(Fmt(runs[2]).ContainsKey("bold"));
        Assert.Equal(true, Fmt(runs[3])["bold"]);
        Assert.False(Fmt(runs[4]).ContainsKey("bold"));
        Assert.Equal(2, handler.LastFindMatchCount);
    }
}
