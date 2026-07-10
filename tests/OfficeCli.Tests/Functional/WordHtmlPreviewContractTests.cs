using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordHtmlPreviewContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void ViewAsHtml_RendersParagraphAndTableStructure()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Preview paragraph" });
        handler.Add("/body", "table", null, new() { ["data"] = "A,B" });

        var html = handler.ViewAsHtml();

        Assert.Contains("<!DOCTYPE html>", html);
        Assert.Contains("Preview paragraph", html);
        Assert.Contains("<table", html);
        Assert.Contains(">A<", html);
        Assert.Contains(">B<", html);
    }

    [Fact]
    public void ViewAsHtml_RendersPictureDataUriAndAltText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "preview.png",
            ["alt"] = "Preview image"
        });

        var html = handler.ViewAsHtml();

        Assert.Contains("<img ", html);
        Assert.Contains("src=\"data:image/png;base64,", html);
        Assert.Contains("alt=\"Preview image\"", html);
    }

    [Fact]
    public void ViewAsHtml_RendersComplexObjectsWithStableAnchorAndObjectMarkers()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Complex preview" });
        handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Preview chart",
            ["categories"] = "Q1,Q2",
            ["data"] = "Revenue:10,20"
        });
        handler.Add("/body", "shape", null, new()
        {
            ["geometry"] = "roundRect",
            ["width"] = "4cm",
            ["height"] = "2cm",
            ["fill"] = "EAF2FF",
            ["alt"] = "Preview shape"
        });
        handler.Add("/body", "textbox", null, new()
        {
            ["text"] = "Preview textbox",
            ["width"] = "4cm",
            ["height"] = "2cm"
        });
        handler.Add("/body", "watermark", null, new() { ["text"] = "DRAFT" });

        var html = handler.ViewAsHtml();

        Assert.Contains("data-path=\"/body/p[1]\"", html);
        Assert.Contains("<svg", html);
        Assert.Contains("Preview chart", html);
        Assert.Contains("background-color:#EAF2FF", html);
        Assert.Contains("Preview textbox", html);
        Assert.Contains("vml-watermark", html);
        Assert.Contains("DRAFT", html);
    }
}
