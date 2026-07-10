using DocumentFormat.OpenXml.Packaging;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class WordValidationAndRawXmlTests : WordTestBase
{
    private const string WordNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string RelationshipsNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Fact]
    public void Validate_HealthyDocumentReturnsNoDiagnostics()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "valid" });

        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void Validate_InvalidChildReturnsStructuredSchemaDiagnostic()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "invalid" });
        handler.RawSet(
            "/document",
            "//w:body/w:p[1]",
            "append",
            $"<w:notAWordChild xmlns:w=\"{WordNamespace}\" />");

        var error = Assert.Single(handler.Validate(), item => item.ErrorType == "Schema");

        Assert.False(string.IsNullOrWhiteSpace(error.Description));
        Assert.Contains("notAWordChild", error.Description, StringComparison.Ordinal);
        Assert.Contains("w:p", error.Path ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("document.xml", error.Part ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void Validate_ReportsDanglingHeaderReferenceWithActionableLocation()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.RawSet(
            "/document",
            "//w:body/w:sectPr",
            "append",
            $"<w:headerReference xmlns:w=\"{WordNamespace}\" xmlns:r=\"{RelationshipsNamespace}\" w:type=\"default\" r:id=\"rIdMissing\" />");

        var error = Assert.Single(handler.Validate(), item => item.ErrorType == "OrphanedReference");

        Assert.Contains("rIdMissing", error.Description, StringComparison.Ordinal);
        Assert.Equal("/w:document/w:body", error.Path);
        Assert.Contains("document.xml", error.Part ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void Validate_RepeatedCallsKeepDiagnosticOrderAndShapeStable()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "invalid" });
        handler.RawSet(
            "/document",
            "//w:body/w:p[1]",
            "append",
            $"<w:notAWordChild xmlns:w=\"{WordNamespace}\" />");

        static object?[] Shape(ValidationError item) =>
        [item.ErrorType, item.Description, item.Path, item.Part];

        var first = handler.Validate().Select(Shape).ToArray();
        var second = handler.Validate().Select(Shape).ToArray();

        Assert.NotEmpty(first);
        Assert.Equal(first, second);
        Assert.All(first, item =>
        {
            Assert.NotNull(item[0]);
            Assert.NotNull(item[1]);
        });
    }

    [Theory]
    [InlineData("/word/document.xml", true)]
    [InlineData("word/document.xml?mode=raw#body", true)]
    [InlineData("/word/_rels/document.xml.rels", true)]
    [InlineData("/[Content_Types].xml", true)]
    [InlineData("/document", false)]
    [InlineData("/word/document.docx", false)]
    [InlineData("", false)]
    public void IsZipUriPath_RecognizesXmlRelsAndManifestWithUriSuffixes(
        string path,
        bool expected)
    {
        Assert.Equal(expected, RawXmlHelper.IsZipUriPath(path));
    }

    [Fact]
    public void TryReadAndFindZipUri_IgnoreQueryFragmentAndLeadingSlash()
    {
        var path = CreateBlankDocx();

        using var document = WordprocessingDocument.Open(path, false);
        var typedPart = RawXmlHelper.FindPartByZipUri(
            document,
            "word/document.xml?cache=false#body");
        var xml = RawXmlHelper.TryReadByZipUri(
            document,
            path,
            "word/document.xml?cache=false#body");

        Assert.Same(document.MainDocumentPart, typedPart);
        Assert.NotNull(xml);
        Assert.StartsWith("<w:document", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<?xml", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void RawZipUri_WithNamespaceXPathReadsAndWritesSamePart()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "zip uri" });

        var before = handler.Raw("/word/document.xml?mode=raw#body");
        handler.RawSet(
            "word/document.xml?mode=raw#body",
            "//w:body/w:p[1]",
            "setattr",
            "w:rsidR=00112233");
        var after = handler.Raw("/word/document.xml#body");

        Assert.Contains("zip uri", before, StringComparison.Ordinal);
        Assert.Contains("00112233", after, StringComparison.Ordinal);
        Assert.Empty(handler.Validate());
    }
}
