using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordStableIdTests : WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void EditableOpen_NormalizesMissingAndDuplicateStableIds()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "one" });
            handler.Add("/body", "paragraph", null, new() { ["text"] = "two" });
            handler.Add("/body", "picture", null, new() { ["src"] = TinyPngDataUri });
            handler.Add("/body", "picture", null, new() { ["src"] = TinyPngDataUri });
        }
        CorruptStableIds(path);

        using (new WordHandler(path, editable: true)) { }

        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Descendants<Paragraph>().ToList();
        Assert.All(paragraphs, p => Assert.False(string.IsNullOrEmpty(p.ParagraphId?.Value)));
        Assert.All(paragraphs, p => Assert.False(string.IsNullOrEmpty(p.TextId?.Value)));
        AssertUnique(paragraphs.Select(p => p.ParagraphId!.Value!));

        var docPropIds = body.Descendants<DW.DocProperties>().Select(dp => dp.Id?.Value).ToList();
        Assert.All(docPropIds, id => Assert.True(id.HasValue));
        AssertUnique(docPropIds.Select(id => id!.Value.ToString()));

        var sdtIds = body.Descendants<SdtId>().Select(id => id.Val?.Value).ToList();
        Assert.All(sdtIds, id => Assert.True(id.HasValue));
        AssertUnique(sdtIds.Select(id => id!.Value.ToString()));

        var bookmarkPairs = body.Descendants<BookmarkStart>()
            .Select(start => new
            {
                StartId = start.Id!.Value!,
                EndId = body.Descendants<BookmarkEnd>().Single(end => end.Id == start.Id).Id!.Value!
            })
            .ToList();
        AssertUnique(bookmarkPairs.Select(p => p.StartId));
        Assert.All(bookmarkPairs, p => Assert.Equal(p.StartId, p.EndId));
    }

    [Fact]
    public void ReadOnlyOpen_DoesNotRewriteStableIds()
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
            handler.Add("/body", "paragraph", null, new() { ["text"] = "one" });
        CorruptStableIds(path);
        var before = File.ReadAllBytes(path);

        using (new WordHandler(path, editable: false)) { }

        Assert.Equal(before, File.ReadAllBytes(path));
    }

    private static void CorruptStableIds(string path)
    {
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document!.Body!;

        foreach (var p in body.Descendants<Paragraph>())
        {
            p.ParagraphId = "00000001";
            p.TextId = null;
        }
        foreach (var dp in body.Descendants<DW.DocProperties>())
            dp.Id = 1U;

        var sectPr = body.GetFirstChild<SectionProperties>();
        body.InsertBefore(new SdtBlock(
            new SdtProperties(new SdtId { Val = 7 }),
            new SdtContentBlock(new Paragraph(new Run(new Text("sdt one"))))), sectPr);
        body.InsertBefore(new SdtBlock(
            new SdtProperties(new SdtId { Val = 7 }),
            new SdtContentBlock(new Paragraph(new Run(new Text("sdt two"))))), sectPr);
        body.InsertBefore(new Paragraph(
            new BookmarkStart { Id = "1", Name = "bm_one" },
            new Run(new Text("bookmark one")),
            new BookmarkEnd { Id = "1" },
            new BookmarkStart { Id = "1", Name = "bm_two" },
            new Run(new Text("bookmark two")),
            new BookmarkEnd { Id = "1" }), sectPr);

        doc.Save();
    }

    private static void AssertUnique(IEnumerable<string> values)
    {
        var list = values.ToList();
        Assert.Equal(list.Count, list.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }
}
