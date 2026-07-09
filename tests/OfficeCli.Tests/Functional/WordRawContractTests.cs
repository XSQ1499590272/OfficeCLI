using DocumentFormat.OpenXml.Packaging;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordRawContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void Raw_ReadsDocumentStylesAndHeaderParts()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/", "header", null, new() { ["text"] = "Raw header" });

        Assert.Contains("<w:document", handler.Raw("/document"));
        Assert.Contains("<w:styles", handler.Raw("/styles"));
        Assert.Contains("Raw header", handler.Raw("/header[1]"));
    }

    [Fact]
    public void RawSet_SetsParagraphAttributeAndValidates()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Raw target" });
        handler.RawSet("/document", "//w:p[1]", "setattr", "w:rsidR=00112233");

        var raw = handler.Raw("/document");

        Assert.Contains("00112233", raw);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedRawSet_UnknownPartDoesNotMutateDocument()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var before = handler.Raw("/document");

        Assert.Throws<ArgumentException>(() =>
            handler.RawSet("/missing", "/w:document", "setattr", "w:background=FFEEAA"));

        Assert.Equal(before, handler.Raw("/document"));
    }

    [Fact]
    public void FailedRawSet_UnmatchedXPathDoesNotMutateDocument()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Raw target" });
        var before = handler.Raw("/document");

        var ex = Assert.Throws<ArgumentException>(() =>
            handler.RawSet("/document", "//w:p[99]", "setattr", "w:rsidR=00112233"));

        Assert.Contains("XPath matched no elements", ex.Message);
        Assert.Equal(before, handler.Raw("/document"));
    }

    [Fact]
    public void AddPart_ChartCreatesReadablePartAndMainDocumentRelationship()
    {
        var path = CreateBlankDocx();
        string relId;

        using (var handler = new WordHandler(path, editable: true))
        {
            (relId, var partPath) = handler.AddPart("/", "chart");

            Assert.False(string.IsNullOrWhiteSpace(relId));
            Assert.Equal("/chart[1]", partPath);
            Assert.Contains("<c:chartSpace", handler.Raw(partPath));
        }

        using var doc = WordprocessingDocument.Open(path, false);
        Assert.IsType<ChartPart>(doc.MainDocumentPart!.GetPartById(relId));
    }

    [Fact]
    public void AddPart_RepeatedChartCreatesDistinctRelationshipsAndPaths()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.AddPart("/", "chart");
        var second = handler.AddPart("/", "chart");

        Assert.NotEqual(first.RelId, second.RelId);
        Assert.Equal("/chart[1]", first.PartPath);
        Assert.Equal("/chart[2]", second.PartPath);
        Assert.Contains("<c:chartSpace", handler.Raw(first.PartPath));
        Assert.Contains("<c:chartSpace", handler.Raw(second.PartPath));
    }

    [Fact]
    public void FailedRawSet_AfterAddPartKeepsExistingRelationshipValid()
    {
        var path = CreateBlankDocx();
        string relId;

        using (var handler = new WordHandler(path, editable: true))
        {
            (relId, _) = handler.AddPart("/", "chart");
            var beforeErrors = handler.Validate();

            var ex = Assert.Throws<ArgumentException>(() =>
                handler.RawSet("/document", "//w:p[99]", "setattr", "w:rsidR=00112233"));
            var afterErrors = handler.Validate();

            Assert.Contains("XPath matched no elements", ex.Message);
            Assert.Equal(beforeErrors.Count, afterErrors.Count);
            Assert.Equal(
                beforeErrors.Select(err => (err.ErrorType, err.Description, err.Path, err.Part)),
                afterErrors.Select(err => (err.ErrorType, err.Description, err.Path, err.Part)));
        }

        using var doc = WordprocessingDocument.Open(path, false);
        Assert.IsType<ChartPart>(doc.MainDocumentPart!.GetPartById(relId));
    }
}
