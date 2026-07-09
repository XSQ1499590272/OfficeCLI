using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordMediaQueryTests : WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void QueryPicture_ReadsInlinePictureMetadata()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "inline.png",
            ["alt"] = "Inline image",
            ["width"] = "1cm",
            ["height"] = "1cm"
        });

        var picture = Assert.Single(handler.Query("picture"));

        Assert.Equal("picture", picture.Type);
        Assert.Equal("inline", Fmt(picture)["wrap"]);
        Assert.Equal("inline.png", Fmt(picture)["name"]);
        Assert.Equal("Inline image", Fmt(picture)["alt"]);
        Assert.Equal("1.0cm", Fmt(picture)["width"]);
        Assert.Equal("1.0cm", Fmt(picture)["height"]);
    }

    [Fact]
    public void QueryPicture_ReadsFloatingAnchorPosition()
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
            ["vRelative"] = "paragraph"
        });

        var picture = Assert.Single(handler.Query("picture"));

        Assert.Equal("square", Fmt(picture)["wrap"]);
        Assert.Equal("2.0cm", Fmt(picture)["hPosition"]);
        Assert.Equal("1.0cm", Fmt(picture)["vPosition"]);
        Assert.Equal("page", Fmt(picture)["hRelative"]);
        Assert.Equal("paragraph", Fmt(picture)["vRelative"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void QueryOle_DoesNotPollutePictureResults()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "real-picture.png"
        });
        handler.Add("/body", "ole", null, new()
        {
            ["src"] = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,SGVsbG8=",
            ["oleKind"] = "package",
            ["contentType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ["embedExt"] = "docx",
            ["progId"] = "Word.Document.12",
            ["display"] = "icon",
            ["name"] = "Embedded Word",
            ["width"] = "2cm",
            ["height"] = "1cm"
        });

        var picture = Assert.Single(handler.Query("picture"));
        var ole = Assert.Single(handler.Query("ole"));

        Assert.Equal("real-picture.png", Fmt(picture)["name"]);
        Assert.Equal("ole", ole.Type);
        Assert.Equal("Word.Document.12", Fmt(ole)["progId"]);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", Fmt(ole)["contentType"]);
        Assert.Equal(5L, Fmt(ole)["fileSize"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddOle_LegacyObjectUsesEmbeddedObjectPartAndReadsBackMetadata()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "ole", null, new()
            {
                ["src"] = "data:application/vnd.ms-excel;base64,TEVHQUNZ",
                ["oleKind"] = "object",
                ["contentType"] = "application/vnd.ms-excel",
                ["progId"] = "Excel.Sheet.8",
                ["display"] = "content",
                ["name"] = "Legacy Excel",
                ["width"] = "3cm",
                ["height"] = "2cm"
            });

            var ole = Assert.Single(handler.Query("ole"));

            Assert.Equal("Excel.Sheet.8", Fmt(ole)["progId"]);
            Assert.Equal("application/vnd.ms-excel", Fmt(ole)["contentType"]);
            Assert.Equal(6L, Fmt(ole)["fileSize"]);
            Assert.Equal("3cm", Fmt(ole)["width"]);
            Assert.Equal("2cm", Fmt(ole)["height"]);
            Assert.Empty(handler.Validate());
        }

        using var document = WordprocessingDocument.Open(path, false);
        var mainPart = document.MainDocumentPart!;
        Assert.Empty(mainPart.EmbeddedPackageParts);
        Assert.Single(mainPart.EmbeddedObjectParts);
    }

    [Fact]
    public void InlinedParts_AddsTopLevelChartPartAndRewritesHostRelId()
    {
        var path = CreateBlankDocx();
        var chartXml = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:lang val="en-US"/>
              <c:chart>
                <c:plotArea/>
              </c:chart>
              <c:externalData r:id="rWorkbook"/>
            </c:chartSpace>
            """));
        var chartStyleXml = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"/>
            """));

        using (var handler = new WordHandler(path, editable: true))
        {
            var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });

            var runPath = handler.Add(paragraphPath, "inlinedparts", null, new()
            {
                ["runXml"] = """
                    <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                         xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                         xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <w:drawing>
                        <wp:inline>
                          <wp:extent cx="914400" cy="914400"/>
                          <wp:docPr id="1" name="Inlined chart"/>
                          <a:graphic>
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                              <c:chart r:id="rOld"/>
                            </a:graphicData>
                          </a:graphic>
                        </wp:inline>
                      </w:drawing>
                    </w:r>
                    """,
                ["part1.relId"] = "rOld",
                ["part1.data"] = $"data:application/vnd.openxmlformats-officedocument.drawingml.chart+xml;base64,{chartXml}",
                ["part1.ext1.relId"] = "rWorkbook",
                ["part1.ext1.type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                ["part1.ext1.target"] = "embeddings/oleObject1.bin",
                ["part1.child1.relId"] = "rStyle",
                ["part1.child1.data"] = $"data:application/vnd.ms-office.chartstyle+xml;base64,{chartStyleXml}"
            });

            Assert.Contains("/r[", runPath);
        }

        using var document = WordprocessingDocument.Open(path, false);
        var mainPart = document.MainDocumentPart!;
        var documentRoot = mainPart.Document ?? throw new InvalidOperationException("Document root not found.");
        var body = documentRoot.Body ?? throw new InvalidOperationException("Document body not found.");
        var chartRef = Assert.Single(body.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>());
        var rewrittenRelId = chartRef.Id!.Value!;

        Assert.NotEqual("rOld", rewrittenRelId);
        Assert.DoesNotContain("\"rOld\"", body.OuterXml);

        var chartPart = Assert.IsType<ChartPart>(mainPart.GetPartById(rewrittenRelId));
        var chartExternalRel = Assert.Single(chartPart.ExternalRelationships);
        Assert.Equal("rWorkbook", chartExternalRel.Id);
        Assert.Equal("embeddings/oleObject1.bin", chartExternalRel.Uri.ToString());
        var stylePartPair = Assert.Single(chartPart.Parts, part => part.RelationshipId == "rStyle");
        Assert.IsType<ChartStylePart>(stylePartPair.OpenXmlPart);
        using var reader = new StreamReader(chartPart.GetStream());
        var chartPartXml = reader.ReadToEnd();
        Assert.Contains("<c:lang val=\"en-US\"", chartPartXml);
        Assert.Contains("r:id=\"rWorkbook\"", chartPartXml);
    }

    [Fact]
    public void InlinedParts_RecreatesHostExternalRelationshipAndRewritesRelId()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });

            handler.Add(paragraphPath, "inlinedparts", null, new()
            {
                ["runXml"] = """
                    <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                         xmlns:v="urn:schemas-microsoft-com:vml">
                      <w:pict>
                        <v:shape>
                          <v:textbox>
                            <w:txbxContent>
                              <w:p>
                                <w:hyperlink r:id="rExt">
                                  <w:r><w:t>External link</w:t></w:r>
                                </w:hyperlink>
                              </w:p>
                            </w:txbxContent>
                          </v:textbox>
                        </v:shape>
                      </w:pict>
                    </w:r>
                    """,
                ["ext1.relId"] = "rExt",
                ["ext1.type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                ["ext1.target"] = "https://example.com/report"
            });
        }

        using var document = WordprocessingDocument.Open(path, false);
        var mainPart = document.MainDocumentPart!;
        var documentRoot = mainPart.Document ?? throw new InvalidOperationException("Document root not found.");
        var body = documentRoot.Body ?? throw new InvalidOperationException("Document body not found.");
        var hyperlinkRel = Assert.Single(mainPart.HyperlinkRelationships);

        Assert.Equal("https://example.com/report", hyperlinkRel.Uri.ToString());
        Assert.NotEqual("rExt", hyperlinkRel.Id);
        Assert.DoesNotContain("\"rExt\"", body.OuterXml);
        Assert.Contains($"\"{hyperlinkRel.Id}\"", body.OuterXml);
    }
}
