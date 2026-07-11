using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordActiveXCarrierRoundTripTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void ActiveXCarrier_DumpBatchReplayRewritesControlAndBinaryRelationships()
    {
        var sourcePath = CreateBlankDocx();
        var targetPath = CreateBlankDocx();

        using (var source = new WordHandler(sourcePath, editable: true))
        {
            var paragraphPath = source.Add("/body", "paragraph", null, new() { ["text"] = "" });
            source.Add(paragraphPath, "inlinedparts", null, new()
            {
                ["runXml"] = """
                    <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:object>
                        <w:control r:id="rControlSource"/>
                      </w:object>
                    </w:r>
                    """,
                ["part1.relId"] = "rControlSource",
                ["part1.data"] = DataUri(
                    "application/vnd.ms-office.activeX+xml",
                    "<ax:ocx xmlns:ax=\"http://schemas.microsoft.com/office/activex/activexml/2010\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rDataSource\" />"),
                ["part1.child1.relId"] = "rDataSource",
                ["part1.child1.data"] = DataUri("application/vnd.ms-office.activeX", "active-x-binary"),
            });

            var items = WordBatchEmitter.EmitWord(source);
            var carrier = Assert.Single(items, item => item.Type == "inlinedparts");
            Assert.Contains("<w:control", carrier.Props!["runXml"], StringComparison.Ordinal);
            Assert.Contains("part1.child1.data", carrier.Props.Keys);

            using var target = new WordHandler(targetPath, editable: true);
            var output = BatchExecutor.ExecuteBatch(
                target,
                JsonSerializer.Serialize(items),
                json: false);
            Assert.Contains("0 failed", output);
        }

        using var document = WordprocessingDocument.Open(targetPath, false);
        var mainPart = document.MainDocumentPart!;
        var bodyXml = mainPart.Document!.Body!.OuterXml;
        Assert.Contains("<w:control", bodyXml, StringComparison.Ordinal);
        Assert.DoesNotContain("rControlSource", bodyXml, StringComparison.Ordinal);

        var controlPart = Assert.Single(mainPart.GetPartsOfType<EmbeddedControlPersistencePart>());
        Assert.Equal("application/vnd.ms-office.activeX+xml", controlPart.ContentType);
        Assert.Single(controlPart.GetPartsOfType<EmbeddedControlPersistenceBinaryDataPart>());
        Assert.Empty(new WordHandler(targetPath, editable: false).Validate());
    }

    private static string DataUri(string contentType, string value)
        => $"data:{contentType};base64,{Convert.ToBase64String(Encoding.UTF8.GetBytes(value))}";
}
