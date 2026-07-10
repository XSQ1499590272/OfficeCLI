using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public class WordDumpBatchRoundTripTests : OfficeCli.Tests.Unit.WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void DumpFullDocument_BatchReplayKeepsParagraphTablePictureAndValidates()
    {
        var sourcePath = CreateBlankDocx();
        var targetPath = CreateBlankDocx();
        string targetPictureRelId;

        using (var source = new WordHandler(sourcePath, editable: true))
        {
            source.Add("/body", "paragraph", null, new() { ["text"] = "Round trip paragraph" });
            source.Add("/body", "table", null, new()
            {
                ["data"] = "A,B;C,D",
                ["colWidths"] = "1200,1800"
            });
            source.Add("/body", "picture", null, new()
            {
                ["src"] = TinyPngDataUri,
                ["name"] = "round-trip.png",
                ["alt"] = "Round trip image",
                ["width"] = "1cm",
                ["height"] = "1cm"
            });

            var dumpedItems = WordBatchEmitter.EmitWord(source);
            var dumpedJson = JsonSerializer.Serialize(dumpedItems);

            using var target = new WordHandler(targetPath, editable: true);
            var output = BatchExecutor.ExecuteBatch(target, dumpedJson, json: false);

            Assert.Contains("0 failed", output);
            Assert.Equal("Round trip paragraph", target.Get("/body/p[1]").Text);

            var table = target.Get("/body/tbl[1]", depth: 2);
            Assert.Equal("table", table.Type);
            Assert.Equal("A", table.Children[0].Children[0].Text);
            Assert.Equal("D", table.Children[1].Children[1].Text);

            var picture = Assert.Single(target.Query("picture"));
            Assert.Equal("round-trip.png", Fmt(picture)["name"]);
            Assert.Equal("Round trip image", Fmt(picture)["alt"]);
            targetPictureRelId = Fmt(picture)["relId"]!.ToString()!;
            Assert.Empty(target.Validate());
        }

        using var doc = WordprocessingDocument.Open(targetPath, false);
        var imagePart = Assert.IsType<ImagePart>(doc.MainDocumentPart!.GetPartById(targetPictureRelId));
        using var stream = imagePart.GetStream();
        using var bytes = new MemoryStream();
        stream.CopyTo(bytes);
        var expected = Convert.FromBase64String(TinyPngDataUri[(TinyPngDataUri.IndexOf(',') + 1)..]);
        Assert.Equal(expected, bytes.ToArray());
    }

    [Fact]
    public void DumpTableSubtree_BatchReplayKeepsOnlyRequestedTable()
    {
        var sourcePath = CreateBlankDocx();
        var targetPath = CreateBlankDocx();

        using (var source = new WordHandler(sourcePath, editable: true))
        {
            source.Add("/body", "paragraph", null, new() { ["text"] = "Before table" });
            source.Add("/body", "table", null, new() { ["data"] = "H1,H2;A,B" });
            source.Add("/body", "paragraph", null, new() { ["text"] = "After table" });

            var dumpedItems = WordBatchEmitter.EmitWord(source, "/body/tbl[1]");
            var dumpedJson = JsonSerializer.Serialize(dumpedItems);

            using var target = new WordHandler(targetPath, editable: true);
            var output = BatchExecutor.ExecuteBatch(target, dumpedJson, json: false);

            Assert.Contains("0 failed", output);
            Assert.Empty(target.Query("paragraph:contains(\"Before table\")"));
            Assert.Empty(target.Query("paragraph:contains(\"After table\")"));

            var table = target.Get("/body/tbl[1]", depth: 2);
            Assert.Equal("H1", table.Children[0].Children[0].Text);
            Assert.Equal("B", table.Children[1].Children[1].Text);
            Assert.Empty(target.Validate());
        }
    }

    [Fact]
    public void DumpBody_EmitsReplayableParagraphTableAndPictureItems()
    {
        var sourcePath = CreateBlankDocx();

        using var source = new WordHandler(sourcePath, editable: true);
        source.Add("/body", "paragraph", null, new() { ["text"] = "Dump paragraph" });
        source.Add("/body", "table", null, new() { ["data"] = "A,B;C,D" });
        source.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "dump.png",
            ["alt"] = "Dump image"
        });

        var items = WordBatchEmitter.EmitWord(source, "/body");

        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Parent == "/body"
            && item.Type == "p"
            && item.Props?["text"] == "Dump paragraph");

        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Parent == "/body"
            && item.Type == "table"
            && item.Props?["rows"] == "2"
            && item.Props?["cols"] == "2");

        Assert.Contains(items, item =>
            item.Command == "set"
            && item.Path == "/body/tbl[last()]/tr[2]/tc[2]/p[last()]"
            && item.Props?["text"] == "D");

        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Type == "picture"
            && item.Parent == "/body/p[last()]"
            && item.Props?["name"] == "dump.png"
            && item.Props?["alt"] == "Dump image"
            && item.Props?["src"].StartsWith("data:image/png;base64,", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void DumpFullDocument_BatchReplayKeepsOlePackage()
    {
        var sourcePath = CreateBlankDocx();
        var targetPath = CreateBlankDocx();

        using (var source = new WordHandler(sourcePath, editable: true))
        {
            source.Add("/body", "ole", null, new()
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

            var dumpedJson = JsonSerializer.Serialize(WordBatchEmitter.EmitWord(source));

            using var target = new WordHandler(targetPath, editable: true);
            var output = BatchExecutor.ExecuteBatch(target, dumpedJson, json: false);
            var ole = Assert.Single(target.Query("ole"));
            var fmt = Fmt(ole);

            Assert.Contains("0 failed", output);
            Assert.Equal("Word.Document.12", fmt["progId"]);
            Assert.Equal("Embedded Word", fmt["name"]);
            Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", fmt["contentType"]);
            Assert.Equal(5L, fmt["fileSize"]);
            Assert.Empty(target.Validate());
        }
    }

    [Fact]
    public void DumpFullDocument_BatchReplayKeepsChartSeries()
    {
        var sourcePath = CreateBlankDocx();
        var targetPath = CreateBlankDocx();

        using (var source = new WordHandler(sourcePath, editable: true))
        {
            source.Add("/body", "chart", null, new()
            {
                ["chartType"] = "column",
                ["title"] = "Round trip chart",
                ["categories"] = "Q1,Q2,Q3",
                ["data"] = "Revenue:10,20,30"
            });

            var dumpedJson = JsonSerializer.Serialize(WordBatchEmitter.EmitWord(source));

            using var target = new WordHandler(targetPath, editable: true);
            var output = BatchExecutor.ExecuteBatch(target, dumpedJson, json: false);
            var chart = Assert.Single(target.Query("chart"));
            var series = target.Get($"{chart.Path}/series[1]");

            Assert.Contains("0 failed", output);
            Assert.Equal("Round trip chart", Fmt(chart)["title"]);
            Assert.Equal("Q1,Q2,Q3", Fmt(chart)["categories"]);
            Assert.Equal("Revenue:10,20,30", Fmt(chart)["series1"]);
            Assert.Equal("10,20,30", Fmt(series)["values"]);
            Assert.Empty(target.Validate());
        }
    }

    [Fact]
    public void DumpFullDocument_BatchReplayKeepsHeaderPictureRelationship()
    {
        var sourcePath = CreateBlankDocx();
        var targetPath = CreateBlankDocx();

        using (var source = new WordHandler(sourcePath, editable: true))
        {
            var headerPath = source.Add("/", "header", null, new() { ["text"] = "" });
            source.Add(headerPath, "picture", null, new()
            {
                ["src"] = TinyPngDataUri,
                ["name"] = "header.png",
                ["alt"] = "Header image"
            });
            var footerPath = source.Add("/", "footer", null, new() { ["text"] = "" });
            source.Add(footerPath, "picture", null, new()
            {
                ["src"] = TinyPngDataUri,
                ["name"] = "footer.png",
                ["alt"] = "Footer image"
            });

            var dumpedJson = JsonSerializer.Serialize(WordBatchEmitter.EmitWord(source));

            using var target = new WordHandler(targetPath, editable: true);
            var output = BatchExecutor.ExecuteBatch(target, dumpedJson, json: false);
            var header = target.Get("/header[1]", depth: 5);
            var picture = Assert.Single(header.Children.SelectMany(p => p.Children), child => child.Type == "picture");
            var footer = target.Get("/footer[1]", depth: 5);
            var footerPicture = Assert.Single(footer.Children.SelectMany(p => p.Children), child => child.Type == "picture");

            Assert.Contains("0 failed", output);
            Assert.Equal("header", header.Type);
            Assert.Equal("header.png", Fmt(picture)["name"]);
            Assert.Equal("Header image", Fmt(picture)["alt"]);
            Assert.Equal("footer", footer.Type);
            Assert.Equal("footer.png", Fmt(footerPicture)["name"]);
            Assert.Equal("Footer image", Fmt(footerPicture)["alt"]);
            Assert.Empty(target.Validate());
        }
    }
}
