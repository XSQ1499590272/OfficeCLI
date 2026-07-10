using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public class WordComplexDumpCarrierRoundTripTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void DumpBatch_ChartCarrierKeepsChildStyleAndExternalRelationship()
    {
        var sourcePath = CreateBlankDocx();
        var targetPath = CreateBlankDocx();
        var chartXml = Convert.ToBase64String(Encoding.UTF8.GetBytes("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:lang val="en-US"/>
              <c:chart><c:plotArea/></c:chart>
              <c:externalData r:id="rWorkbook"/>
            </c:chartSpace>
            """));
        var styleXml = Convert.ToBase64String(Encoding.UTF8.GetBytes("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"/>
            """));

        using (var source = new WordHandler(sourcePath, editable: true))
        {
            var paragraphPath = source.Add("/body", "paragraph", null, new() { ["text"] = "" });
            source.Add(paragraphPath, "inlinedparts", null, new()
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
                          <wp:docPr id="1" name="Dumped carrier chart"/>
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
                ["part1.child1.relId"] = "rStyle",
                ["part1.child1.data"] = $"data:application/vnd.ms-office.chartstyle+xml;base64,{styleXml}",
                ["part1.ext1.relId"] = "rWorkbook",
                ["part1.ext1.type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                ["part1.ext1.target"] = "embeddings/oleObject1.bin"
            });

            var batch = JsonSerializer.Serialize(WordBatchEmitter.EmitWord(source));
            using var target = new WordHandler(targetPath, editable: true);
            var output = BatchExecutor.ExecuteBatch(target, batch, json: false);

            Assert.Contains("0 failed", output);
        }

        using var document = WordprocessingDocument.Open(targetPath, false);
        var mainPart = document.MainDocumentPart!;
        var body = mainPart.Document!.Body!;
        var chartReference = Assert.Single(body.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>());
        var rewrittenHostRelId = chartReference.Id!.Value!;
        Assert.NotEqual("rOld", rewrittenHostRelId);
        Assert.DoesNotContain("\"rOld\"", body.OuterXml);

        var chartPart = Assert.IsType<ChartPart>(mainPart.GetPartById(rewrittenHostRelId));
        var external = Assert.Single(chartPart.ExternalRelationships);
        Assert.Equal("rWorkbook", external.Id);
        Assert.Equal("embeddings/oleObject1.bin", external.Uri.ToString());
        Assert.Contains(chartPart.Parts, part => part.RelationshipId == "rStyle"
            && part.OpenXmlPart is ChartStylePart);

        using var chartReader = new StreamReader(chartPart.GetStream());
        var rebuiltChartXml = chartReader.ReadToEnd();
        Assert.Contains("<c:lang val=\"en-US\"", rebuiltChartXml);
        Assert.Contains("r:id=\"rWorkbook\"", rebuiltChartXml);
    }
}
