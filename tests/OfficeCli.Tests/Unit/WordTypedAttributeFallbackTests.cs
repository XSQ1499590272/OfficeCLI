using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordTypedAttributeFallbackTests
{
    [Fact]
    public void TrySet_AcceptsAliasesAndNormalizesColorAttributes()
    {
        var paragraphProperties = new ParagraphProperties();

        Assert.True(TypedAttributeFallback.TrySet(paragraphProperties, "ind.firstLine", "240"));
        Assert.True(TypedAttributeFallback.TrySet(paragraphProperties, "shading.fill", "#00FF00"));

        var indentation = paragraphProperties.GetFirstChild<Indentation>();
        Assert.NotNull(indentation);
        Assert.Equal("240", indentation!.FirstLine?.Value);

        var shading = paragraphProperties.GetFirstChild<Shading>();
        Assert.NotNull(shading);
        Assert.Equal("00FF00", shading!.Fill?.Value);
        Assert.Equal(ShadingPatternValues.Clear, shading.Val?.Value);
    }

    [Fact]
    public void TrySet_RejectsUnknownPairsAndRequiresExistingNestedStructure()
    {
        var paragraphProperties = new ParagraphProperties();

        Assert.False(TypedAttributeFallback.TrySet(paragraphProperties, "ind.notAnAttribute", "1"));
        Assert.False(TypedAttributeFallback.TrySet(paragraphProperties, "pBdr.top.val", "single"));

        paragraphProperties.AppendChild(new ParagraphBorders(new TopBorder()));
        Assert.True(TypedAttributeFallback.TrySet(paragraphProperties, "pBdr.top.val", "single"));
        Assert.Equal(BorderValues.Single, paragraphProperties.GetFirstChild<ParagraphBorders>()!
            .GetFirstChild<TopBorder>()!.Val!.Value);
    }

    [Fact]
    public void TrySet_ResolvesFontAliasAgainstRunProperties()
    {
        var runProperties = new RunProperties();

        Assert.True(TypedAttributeFallback.TrySet(runProperties, "font.ascii", "Arial"));

        var fonts = runProperties.GetFirstChild<RunFonts>();
        Assert.NotNull(fonts);
        Assert.Equal("Arial", fonts!.Ascii?.Value);
    }

    [Fact]
    public void TrySet_MergesCaseInsensitiveFontAliasesIntoOneRunFontsElement()
    {
        var runProperties = new RunProperties();

        Assert.True(TypedAttributeFallback.TrySet(runProperties, "FONT.ascii", "Arial"));
        Assert.True(TypedAttributeFallback.TrySet(runProperties, "font.hAnsi", "Arial Unicode MS"));

        var fonts = Assert.Single(runProperties.Elements<RunFonts>());
        Assert.Equal("Arial", fonts.Ascii?.Value);
        Assert.Equal("Arial Unicode MS", fonts.HighAnsi?.Value);
    }

    [Fact]
    public void TrySet_RequiresExistingNestedContainerAndPreservesNestedColorValue()
    {
        var paragraphProperties = new ParagraphProperties(
            new ParagraphBorders(new TopBorder()));

        Assert.True(TypedAttributeFallback.TrySet(
            paragraphProperties, "border.top.color", "#112233"));
        Assert.Equal("#112233", paragraphProperties.GetFirstChild<ParagraphBorders>()!
            .GetFirstChild<TopBorder>()!.Color?.Value);

        var missing = new ParagraphProperties();
        Assert.False(TypedAttributeFallback.TrySet(missing, "border.top.color", "#112233"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("unknown")]
    [InlineData(".attr")]
    [InlineData("element.")]
    public void TrySet_RejectsMalformedOrUndottedKeys(string key)
    {
        Assert.False(TypedAttributeFallback.TrySet(new ParagraphProperties(), key, "value"));
    }
}
