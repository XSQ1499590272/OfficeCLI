using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public class WordDocumentFidelityTests : OfficeCli.Tests.Unit.WordTestBase
{
    private const string CustomXmlPayload = "<root><sentinel>preserve-me</sentinel></root>";
    private const string CustomXmlPropertiesPayload =
        "<ds:datastoreItem ds:itemID=\"{11111111-1111-1111-1111-111111111111}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\" />";

    [Fact]
    public void UnrelatedWordMutation_PreservesCustomXmlPackagePart()
    {
        var path = CreateBlankDocx();
        AddCustomXmlPart(path);

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "unrelated mutation" });
        }

        var parts = ReadAuxiliaryParts(path);
        Assert.Contains("<sentinel>preserve-me</sentinel>", parts.CustomXml);
        Assert.Contains("11111111-1111-1111-1111-111111111111", parts.CustomXmlProperties);
        Assert.Contains("PreservedFont", parts.FontTable);
        Assert.Contains("webSettings", parts.WebSettings);
        using var reopened = new WordHandler(path, editable: false);
        Assert.Contains("unrelated mutation", reopened.Query("paragraph").Select(node => node.Text));
        Assert.Empty(reopened.Validate());
    }

    [Fact]
    public void DumpBatchReplay_PreservesCustomXmlDataStoreAndParagraphContent()
    {
        var sourcePath = CreateBlankDocx();
        var targetPath = CreateBlankDocx();
        AddCustomXmlPart(sourcePath);

        using (var source = new WordHandler(sourcePath, editable: true))
        {
            source.Add("/body", "paragraph", null, new() { ["text"] = "source paragraph" });
            var items = WordBatchEmitter.EmitWord(source);
            var json = JsonSerializer.Serialize(items);

            using var target = new WordHandler(targetPath, editable: true);
            var output = BatchExecutor.ExecuteBatch(target, json, json: false);
            Assert.Contains("0 failed", output);
            Assert.Contains("source paragraph", target.Query("paragraph").Select(node => node.Text));
            Assert.Empty(target.Validate());
        }

        var parts = ReadAuxiliaryParts(targetPath);
        Assert.Contains("<sentinel>preserve-me</sentinel>", parts.CustomXml);
        Assert.Contains("11111111-1111-1111-1111-111111111111", parts.CustomXmlProperties);
        Assert.Contains("PreservedFont", parts.FontTable);
        Assert.Contains("webSettings", parts.WebSettings);
    }

    private static void AddCustomXmlPart(string path)
    {
        using var document = WordprocessingDocument.Open(path, true);
        var main = document.MainDocumentPart!;
        var part = main.AddCustomXmlPart(CustomXmlPartType.CustomXml, "rIdCustomPreserved");
        part.FeedData(new MemoryStream(Encoding.UTF8.GetBytes(CustomXmlPayload)));
        var properties = part.AddNewPart<CustomXmlPropertiesPart>("rIdCustomPropertiesPreserved");
        properties.FeedData(new MemoryStream(Encoding.UTF8.GetBytes(CustomXmlPropertiesPayload)));

        var fontTable = main.AddNewPart<FontTablePart>("rIdFontPreserved");
        fontTable.Fonts = new Fonts(new Font { Name = "PreservedFont" });
        fontTable.Fonts.Save();

        var webSettings = main.AddNewPart<WebSettingsPart>("rIdWebPreserved");
        webSettings.WebSettings = new WebSettings();
        webSettings.WebSettings.Save();
    }

    private static (string CustomXml, string CustomXmlProperties, string FontTable, string WebSettings) ReadAuxiliaryParts(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var main = document.MainDocumentPart!;
        var customParts = main.CustomXmlParts.ToList();
        Assert.NotEmpty(customParts);
        var customXml = new StringBuilder();
        var customXmlProperties = new StringBuilder();
        foreach (var part in customParts)
        {
            using var reader = new StreamReader(part.GetStream(), Encoding.UTF8);
            customXml.Append(reader.ReadToEnd());
            if (part.CustomXmlPropertiesPart is { } properties)
            {
                using var propertiesReader = new StreamReader(properties.GetStream(), Encoding.UTF8);
                customXmlProperties.Append(propertiesReader.ReadToEnd());
            }
        }
        return (
            customXml.ToString(),
            customXmlProperties.ToString(),
            main.FontTablePart?.Fonts?.OuterXml ?? "",
            main.WebSettingsPart?.WebSettings?.OuterXml ?? "");
    }
}
