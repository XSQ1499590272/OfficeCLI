using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordWatermarkContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void AddWatermark_ReadsBackTextAndVmlProperties()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var watermarkPath = handler.Add("/body", "watermark", null, new()
        {
            ["text"] = "DRAFT",
            ["color"] = "#C0C0C0",
            ["font"] = "Calibri",
            ["rotation"] = "-45",
            ["opacity"] = ".5",
            ["size"] = "72",
            ["width"] = "415pt",
            ["height"] = "207.5pt"
        });

        var watermark = handler.Get("/watermark");
        var queried = Assert.Single(handler.Query("watermark"));
        var fmt = Fmt(watermark);

        Assert.Equal("/watermark", watermarkPath);
        Assert.Equal("/watermark", queried.Path);
        Assert.Equal("watermark", watermark.Type);
        Assert.Equal("DRAFT", watermark.Text);
        Assert.Equal("DRAFT", fmt["text"]);
        Assert.Equal("#C0C0C0", fmt["color"]);
        Assert.Equal("Calibri", fmt["font"]);
        Assert.Equal("315", fmt["rotation"]);
        Assert.Equal("0.5", fmt["opacity"]);
        Assert.Equal("72pt", fmt["size"]);
        Assert.Equal("415pt", fmt["width"]);
        Assert.Equal("207.5pt", fmt["height"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetWatermark_UpdatesSingletonAndRemoveClearsQueryResult()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "watermark", null, new() { ["text"] = "DRAFT" });

        var unsupported = handler.Set("/watermark", new()
        {
            ["text"] = "CONFIDENTIAL",
            ["color"] = "FF0000",
            ["font"] = "Aptos",
            ["rotation"] = "30",
            ["opacity"] = "0.25",
            ["size"] = "48",
            ["width"] = "300pt",
            ["height"] = "120pt",
            ["unknown"] = "ignored"
        });
        var updated = handler.Get("/watermark");
        var fmt = Fmt(updated);

        Assert.Contains("unknown", unsupported);
        Assert.Equal("CONFIDENTIAL", updated.Text);
        Assert.Equal("#FF0000", fmt["color"]);
        Assert.Equal("Aptos", fmt["font"]);
        Assert.Equal("30", fmt["rotation"]);
        Assert.Equal("0.25", fmt["opacity"]);
        Assert.Equal("48pt", fmt["size"]);
        Assert.Equal("300pt", fmt["width"]);
        Assert.Equal("120pt", fmt["height"]);

        handler.Remove("/watermark");

        Assert.Empty(handler.Query("watermark"));
        Assert.Equal("(no watermark)", handler.Get("/watermark").Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddWatermark_WithInvalidRotationThrowsBeforeMutation()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var ex = Assert.Throws<ArgumentException>(() =>
            handler.Add("/body", "watermark", null, new()
            {
                ["text"] = "DRAFT",
                ["rotation"] = "tilted"
            }));

        Assert.Contains("rotation", ex.Message);
        Assert.Empty(handler.Query("watermark"));
    }

    [Fact]
    public void AddWatermark_WithImageCurrentlyIgnoresImageAndCreatesDefaultTextWatermark()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "watermark", null, new()
        {
            ["image"] = TinyPngDataUri,
            ["width"] = "200pt",
            ["height"] = "100pt"
        });
        var watermark = handler.Get("/watermark");

        Assert.Equal("DRAFT", watermark.Text);
        Assert.Equal("DRAFT", Fmt(watermark)["text"]);
        Assert.Equal("200pt", Fmt(watermark)["width"]);
        Assert.Equal("100pt", Fmt(watermark)["height"]);
        Assert.Empty(handler.Validate());
    }
}
