using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordTocBuilderTests : WordTestBase
{
    [Fact]
    public void RegenerateAllTocs_FiltersHeadingLevelsAndAddsHyperlinksWithoutPageFields()
    {
        var path = CreateBlankDocx();

        using (var doc = WordprocessingDocument.Open(path, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            AppendHeading(body, "Chapter", "Heading1");
            AppendHeading(body, "Section", "Heading2");
            AppendHeading(body, "Appendix", "Heading4");
            AppendToc(body, " TOC \\o \"1-2\" \\h \\n ");

            WordTocBuilder.RegenerateAllTocs(doc);
            doc.MainDocumentPart.Document.Save();
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var paragraphs = reopened.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>().ToList();
        var entries = paragraphs
            .Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value is "TOC1" or "TOC2")
            .ToList();

        Assert.Equal(2, entries.Count);
        Assert.Equal("Chapter", entries[0].InnerText);
        Assert.Equal("Section", entries[1].InnerText);
        Assert.All(entries, entry =>
        {
            var hyperlink = Assert.Single(entry.Elements<Hyperlink>());
            Assert.StartsWith("_Toc", hyperlink.Anchor?.Value);
            Assert.Empty(entry.Descendants<FieldCode>());
            Assert.Empty(entry.Descendants<TabChar>());
        });
    }

    [Fact]
    public void RegenerateAllTocs_DefaultSpecIncludesPageRefAndNoHyperlink()
    {
        var path = CreateBlankDocx();

        using (var doc = WordprocessingDocument.Open(path, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            AppendHeading(body, "Overview", "Heading1");
            AppendToc(body, " TOC ");

            WordTocBuilder.RegenerateAllTocs(doc);
            doc.MainDocumentPart.Document.Save();
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var entry = Assert.Single(reopened.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>(),
            p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TOC1");

        Assert.StartsWith("Overview", entry.InnerText);
        Assert.Empty(entry.Elements<Hyperlink>());
        Assert.Contains(entry.Descendants<FieldCode>(), field =>
            field.Text?.Contains("PAGEREF _Toc", StringComparison.Ordinal) == true);
        Assert.NotEmpty(entry.Descendants<TabChar>());
        Assert.Contains(entry.Descendants<Text>(), text => text.Text == "0");
    }

    [Fact]
    public void RegenerateAllTocs_ReusesExistingHeadingBookmark()
    {
        var path = CreateBlankDocx();

        using (var doc = WordprocessingDocument.Open(path, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var heading = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new BookmarkStart { Id = "7", Name = "_TocExisting" },
                new Run(new Text("Existing heading")),
                new BookmarkEnd { Id = "7" });
            body.Append(heading);
            AppendToc(body, " TOC \\h ");

            WordTocBuilder.RegenerateAllTocs(doc);
            doc.MainDocumentPart.Document.Save();
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var entry = Assert.Single(reopened.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>(),
            p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TOC1");
        var hyperlink = Assert.Single(entry.Elements<Hyperlink>());

        Assert.Equal("_TocExisting", hyperlink.Anchor?.Value);
        Assert.Equal(1, reopened.MainDocumentPart.Document.Body!.Descendants<BookmarkStart>()
            .Count(bookmark => bookmark.Name?.Value == "_TocExisting"));
    }

    [Fact]
    public void RegenerateAllTocs_LeavesDocumentWithoutTocFieldUsable()
    {
        var path = CreateBlankDocx();

        using (var doc = WordprocessingDocument.Open(path, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            AppendHeading(body, "Heading without toc", "Heading1");

            WordTocBuilder.RegenerateAllTocs(doc);
            doc.MainDocumentPart.Document.Save();
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var heading = Assert.Single(reopened.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>(), p => p.InnerText == "Heading without toc");

        Assert.StartsWith("_Toc", heading.Descendants<BookmarkStart>().Single().Name?.Value);
        Assert.Empty(heading.Descendants<FieldCode>());
    }

    private static void AppendHeading(Body body, string text, string styleId)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
            new Run(new Text(text))));
    }

    private static void AppendToc(Body body, string instruction)
    {
        body.Append(new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(instruction)),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("old result")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End })));
    }
}
