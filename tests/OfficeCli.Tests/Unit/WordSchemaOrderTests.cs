using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordSchemaOrderTests
{
    [Theory]
    [InlineData("top-right", "topright")]
    [InlineData("TOP_RIGHT", "topright")]
    [InlineData("top right", "topright")]
    [InlineData(" left ", "left")]
    [InlineData(null, "")]
    public void SchemaKeyNormalizer_UsesCanonicalCaseAndPunctuation(string? raw, string expected)
    {
        Assert.Equal(expected, SchemaKeyNormalizer.Normalize(raw));
    }

    [Fact]
    public void Place_MovesNewParagraphPropertyBeforeLaterSchemaSiblings()
    {
        var properties = new ParagraphProperties();
        var alignment = new Justification { Val = JustificationValues.Center };
        var style = new ParagraphStyleId { Val = "Normal" };
        properties.Append(alignment);
        properties.Append(style);

        SchemaOrder.Place(properties, style);

        Assert.Same(style, properties.ChildElements[0]);
        Assert.Same(alignment, properties.ChildElements[1]);
    }

    [Fact]
    public void Place_IsNoOpForDetachedChildren()
    {
        var properties = new ParagraphProperties();
        var style = new ParagraphStyleId { Val = "Normal" };

        SchemaOrder.Place(properties, style);

        Assert.Empty(properties.ChildElements);
        Assert.Null(style.Parent);
    }
}
