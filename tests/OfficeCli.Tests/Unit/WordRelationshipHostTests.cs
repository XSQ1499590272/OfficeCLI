using DocumentFormat.OpenXml.Packaging;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class WordRelationshipHostTests : WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void PicturesInHeaderAndFooterUseHostRelationshipsWithDistinctIds()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/", "header", null, new() { ["text"] = "header" });
            handler.Add("/", "footer", null, new() { ["text"] = "footer" });

            handler.Add("/header[1]", "picture", null, new()
            {
                ["src"] = TinyPngDataUri,
                ["name"] = "header-one.png"
            });
            handler.Add("/header[1]", "picture", null, new()
            {
                ["src"] = TinyPngDataUri,
                ["name"] = "header-two.png"
            });
            handler.Add("/footer[1]", "picture", null, new()
            {
                ["src"] = TinyPngDataUri,
                ["name"] = "footer-one.png"
            });

            Assert.Contains("<w:drawing", handler.Raw("/header[1]"), StringComparison.Ordinal);
            Assert.Contains("<w:drawing", handler.Raw("/footer[1]"), StringComparison.Ordinal);
            Assert.Empty(handler.Validate());
        }

        using var document = WordprocessingDocument.Open(path, false);
        var main = document.MainDocumentPart!;
        var header = Assert.Single(main.HeaderParts);
        var footer = Assert.Single(main.FooterParts);
        var headerImageIds = header.Parts
            .Where(pair => pair.OpenXmlPart is ImagePart)
            .Select(pair => pair.RelationshipId)
            .ToArray();
        var footerImageIds = footer.Parts
            .Where(pair => pair.OpenXmlPart is ImagePart)
            .Select(pair => pair.RelationshipId)
            .ToArray();

        Assert.Equal(2, headerImageIds.Length);
        Assert.Equal(2, headerImageIds.Distinct(StringComparer.Ordinal).Count());
        Assert.Single(footerImageIds);
        Assert.DoesNotContain(main.Parts, pair => pair.OpenXmlPart is ImagePart);
    }
}
