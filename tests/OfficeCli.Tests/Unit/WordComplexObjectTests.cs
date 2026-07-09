using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordComplexObjectTests : WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void Equation_DisplayAndInlineSetRemove_ReadBackCurrentPathsAndText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var displayPath = handler.Add("/body", "equation", null, new()
        {
            ["formula"] = "x",
            ["mode"] = "display"
        });
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Before " });
        var inlinePath = handler.Add(paragraphPath, "equation", null, new()
        {
            ["formula"] = "a",
            ["mode"] = "inline"
        });

        Assert.Equal("/body/oMathPara[1]", displayPath);
        Assert.EndsWith("/oMath[1]", inlinePath);
        Assert.Equal("display", Fmt(handler.Get(displayPath))["mode"]);
        Assert.Equal("inline", Fmt(handler.Get(inlinePath))["mode"]);

        Assert.Empty(handler.Set(displayPath, new() { ["formula"] = "y" }));
        Assert.Empty(handler.Set(inlinePath, new() { ["formula"] = "b" }));
        Assert.Equal("y", handler.Get(displayPath).Text);
        Assert.Equal("b", handler.Get(inlinePath).Text);

        handler.Remove(displayPath);
        handler.Remove(inlinePath);

        Assert.Empty(handler.Query("equation"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void Textbox_ReadsAddressableTreeAndSetSurfaceKeepsTextAddOnly()
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
            ["line.width"] = "1pt"
        });

        var textbox = handler.Get(textboxPath, depth: 1);
        Assert.Equal("/body/textbox[1]", textboxPath);
        Assert.Equal("txbxContent", textbox.Type);
        Assert.Equal("Sidebar note", textbox.Text);
        Assert.Equal($"{textboxPath}/p[1]", Assert.Single(textbox.Children).Path);
        Assert.Equal("Sidebar note", handler.Get($"{textboxPath}/p[1]/r[1]").Text);

        Assert.Empty(handler.Set(textboxPath, new()
        {
            ["fill"] = "FFECEC",
            ["line.color"] = "C00000",
            ["line.width"] = "2pt",
            ["width"] = "5cm",
            ["height"] = "3cm",
            ["geometry"] = "ellipse"
        }));
        Assert.Contains("text", handler.Set(textboxPath, new() { ["text"] = "Changed text" }));
        Assert.Contains("position keys", Assert.Throws<ArgumentException>(() =>
            handler.Set(textboxPath, new() { ["hPosition"] = "3cm" })).Message);
        Assert.Equal("Sidebar note", handler.Get($"{textboxPath}/p[1]").Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void Shape_ReadsRawDrawingTreeSetAliasesAndRemoveBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var shapePath = handler.Add("/body", "shape", null, new()
        {
            ["geometry"] = "rect",
            ["width"] = "4cm",
            ["height"] = "2cm",
            ["fill"] = "EAF2FF"
        });

        Assert.Equal("/body/shape[1]", shapePath);
        Assert.Equal("wsp", handler.Get(shapePath, depth: 1).Type);

        var unsupported = handler.Set(shapePath, new()
        {
            ["preset"] = "ellipse",
            ["width"] = "5cm",
            ["height"] = "3cm",
            ["fillcolor"] = "FFECEC",
            ["linecolor"] = "C00000",
            ["linewidth"] = "2pt",
            ["hPosition"] = "3cm"
        });
        var geometry = handler.Get($"{shapePath}/spPr[1]/prstGeom[1]");
        var fillColor = handler.Get($"{shapePath}/spPr[1]/solidFill[1]/srgbClr[1]");
        var extents = handler.Get($"{shapePath}/spPr[1]/xfrm[1]/ext[1]");

        Assert.Contains("hPosition", unsupported);
        Assert.Equal("ellipse", Fmt(geometry)["prst"]);
        Assert.Equal("FFECEC", Fmt(fillColor)["val"]);
        Assert.Equal("1800000", Fmt(extents)["cx"]);
        Assert.Equal("1080000", Fmt(extents)["cy"]);

        handler.Remove(shapePath);

        Assert.Throws<ArgumentException>(() => handler.Get(shapePath));
        Assert.DoesNotContain(handler.Get("/body", depth: 1).Children, child => child.Path.Contains("/shape["));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void Watermark_TextImageFallbackInvalidRotationAndRemoveBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        Assert.Throws<ArgumentException>(() =>
            handler.Add("/body", "watermark", null, new()
            {
                ["text"] = "DRAFT",
                ["rotation"] = "tilted"
            }));
        Assert.Empty(handler.Query("watermark"));

        var watermarkPath = handler.Add("/body", "watermark", null, new()
        {
            ["image"] = TinyPngDataUri,
            ["width"] = "200pt",
            ["height"] = "100pt"
        });
        var imageFallback = handler.Get("/watermark");
        Assert.Equal("/watermark", watermarkPath);
        Assert.Equal("DRAFT", imageFallback.Text);
        Assert.Equal("200pt", Fmt(imageFallback)["width"]);
        Assert.Equal("100pt", Fmt(imageFallback)["height"]);

        var unsupported = handler.Set("/watermark", new()
        {
            ["text"] = "CONFIDENTIAL",
            ["color"] = "FF0000",
            ["rotation"] = "30",
            ["opacity"] = "0.25",
            ["unknown"] = "ignored"
        });
        var updated = handler.Get("/watermark");

        Assert.Contains("unknown", unsupported);
        Assert.Equal("CONFIDENTIAL", updated.Text);
        Assert.Equal("#FF0000", Fmt(updated)["color"]);
        Assert.Equal("30", Fmt(updated)["rotation"]);
        Assert.Equal("0.25", Fmt(updated)["opacity"]);

        handler.Remove("/watermark");

        Assert.Empty(handler.Query("watermark"));
        Assert.Equal("(no watermark)", handler.Get("/watermark").Text);
        Assert.Empty(handler.Validate());
    }
}
