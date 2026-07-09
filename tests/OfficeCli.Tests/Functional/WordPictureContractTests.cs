using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordPictureContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void AddPicture_FromDataUri_ReadsBackAltSizeAndInlineWrap()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var hostParagraphPath = handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["alt"] = "Tiny image",
            ["name"] = "tiny.png",
            ["width"] = "1cm",
            ["height"] = "1cm"
        });

        var hostParagraph = handler.Get(hostParagraphPath);
        var picture = Assert.Single(handler.Query("picture"));

        Assert.Equal("paragraph", hostParagraph.Type);
        Assert.Equal("picture", picture.Type);
        Assert.Equal("inline", Fmt(picture)["wrap"]);
        Assert.Equal("Tiny image", Fmt(picture)["alt"]);
        Assert.Equal("tiny.png", Fmt(picture)["name"]);
        Assert.Equal("1.0cm", Fmt(picture)["width"]);
        Assert.Equal("1.0cm", Fmt(picture)["height"]);
    }

    [Fact]
    public void QueryPicture_NoAltPseudoOnlyReturnsImagesWithoutDescription()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "with-alt.png",
            ["alt"] = "Has alternate text"
        });
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "missing-alt.png"
        });

        var result = Assert.Single(handler.Query("picture:no-alt"));

        Assert.Equal("picture", result.Type);
        Assert.Equal("missing-alt.png", Fmt(result)["name"]);
    }

    [Fact]
    public void AddPicture_WithFloatingWrap_ReadsBackAnchorPosition()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "floating.png",
            ["wrap"] = "square",
            ["hPosition"] = "2cm",
            ["vPosition"] = "1cm",
            ["hRelative"] = "page",
            ["vRelative"] = "paragraph",
            ["width"] = "3cm",
            ["height"] = "2cm"
        });

        var picture = Assert.Single(handler.Query("picture"));

        Assert.Equal("square", Fmt(picture)["wrap"]);
        Assert.Equal("floating.png", Fmt(picture)["name"]);
        Assert.Equal("2.0cm", Fmt(picture)["hPosition"]);
        Assert.Equal("1.0cm", Fmt(picture)["vPosition"]);
        Assert.Equal("page", Fmt(picture)["hRelative"]);
        Assert.Equal("paragraph", Fmt(picture)["vRelative"]);
        Assert.Equal("3.0cm", Fmt(picture)["width"]);
        Assert.Equal("2.0cm", Fmt(picture)["height"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddPicture_WithAlignedBehindTextAnchor_ReadsBackAnchorFlags()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "behind.png",
            ["wrap"] = "square",
            ["behindText"] = "true",
            ["hAlign"] = "center",
            ["vAlign"] = "bottom",
            ["hRelative"] = "page",
            ["vRelative"] = "page"
        });

        var picture = Assert.Single(handler.Query("picture"));
        var fmt = Fmt(picture);

        Assert.Equal("square", fmt["wrap"]);
        Assert.Equal(true, fmt["anchor"]);
        Assert.Equal(true, fmt["behindText"]);
        Assert.Equal("center", fmt["hAlign"]);
        Assert.Equal("bottom", fmt["vAlign"]);
        Assert.Equal("page", fmt["hRelative"]);
        Assert.Equal("page", fmt["vRelative"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddPicture_WithCropDecorativeAndLink_ReadsBackCanonicalMetadata()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "linked-cropped.png",
            ["alt"] = "Linked cropped image",
            ["crop"] = "10,20,30,40",
            ["decorative"] = "true",
            ["link"] = "https://example.com/image"
        });

        var picture = Assert.Single(handler.Query("picture"));
        var fmt = Fmt(picture);

        Assert.Equal("linked-cropped.png", fmt["name"]);
        Assert.Equal("Linked cropped image", fmt["alt"]);
        Assert.Equal("10,20,30,40", fmt["crop"]);
        Assert.Equal(true, fmt["decorative"]);
        Assert.Equal("https://example.com/image", fmt["link"]);
        Assert.Empty(handler.Validate());
    }
}
