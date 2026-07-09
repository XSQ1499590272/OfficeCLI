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
}
