using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordSectionRelationshipTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void MultipleSections_KeepHeaderFooterReferencesBoundToTheirOwnParts()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "first section" });
            handler.Add("/body", "section", null, new() { ["type"] = "nextPage" });
            handler.Add("/body", "paragraph", null, new() { ["text"] = "second section" });

            var sections = handler.Query("section");
            Assert.True(sections.Count >= 2);
            var firstSection = sections[0].Path;
            var lastSection = sections[^1].Path;

            handler.Add(firstSection, "header", null, new()
            {
                ["type"] = "default",
                ["text"] = "first header"
            });
            handler.Add(firstSection, "footer", null, new()
            {
                ["type"] = "default",
                ["text"] = "first footer"
            });
            handler.Add(lastSection, "header", null, new()
            {
                ["type"] = "default",
                ["text"] = "last header"
            });
            handler.Add(lastSection, "footer", null, new()
            {
                ["type"] = "default",
                ["text"] = "last footer"
            });
            handler.Add(lastSection, "header", null, new()
            {
                ["type"] = "first",
                ["text"] = "first-page header"
            });
            handler.Add(lastSection, "header", null, new()
            {
                ["type"] = "even",
                ["text"] = "even-page header"
            });
            handler.Add(lastSection, "footer", null, new()
            {
                ["type"] = "first",
                ["text"] = "first-page footer"
            });
            handler.Add(lastSection, "footer", null, new()
            {
                ["type"] = "even",
                ["text"] = "even-page footer"
            });

            Assert.Contains("first header", handler.Query("header").Select(node => node.Text));
            Assert.Contains("last header", handler.Query("header").Select(node => node.Text));
            Assert.Contains("first footer", handler.Query("footer").Select(node => node.Text));
            Assert.Contains("last footer", handler.Query("footer").Select(node => node.Text));
            Assert.Empty(handler.Validate());
        }

        using var document = WordprocessingDocument.Open(path, false);
        var main = document.MainDocumentPart!;
        var sectionProperties = main.Document!.Body!.Descendants<SectionProperties>().ToList();
        Assert.True(sectionProperties.Count >= 2);

        var headerIds = sectionProperties
            .SelectMany(section => section.Elements<HeaderReference>())
            .Select(reference => reference.Id?.Value)
            .Where(id => !string.IsNullOrEmpty(id))
            .ToList();
        var footerIds = sectionProperties
            .SelectMany(section => section.Elements<FooterReference>())
            .Select(reference => reference.Id?.Value)
            .Where(id => !string.IsNullOrEmpty(id))
            .ToList();

        Assert.Equal(4, headerIds.Count);
        Assert.Equal(4, footerIds.Count);
        Assert.All(headerIds, id => Assert.IsType<HeaderPart>(main.GetPartById(id!)));
        Assert.All(footerIds, id => Assert.IsType<FooterPart>(main.GetPartById(id!)));
        Assert.Contains(main.HeaderParts, part => part.Header!.InnerText.Contains("first header"));
        Assert.Contains(main.HeaderParts, part => part.Header!.InnerText.Contains("last header"));
        Assert.Contains(main.HeaderParts, part => part.Header!.InnerText.Contains("first-page header"));
        Assert.Contains(main.HeaderParts, part => part.Header!.InnerText.Contains("even-page header"));
        Assert.Contains(main.FooterParts, part => part.Footer!.InnerText.Contains("first footer"));
        Assert.Contains(main.FooterParts, part => part.Footer!.InnerText.Contains("last footer"));
        Assert.Contains(main.FooterParts, part => part.Footer!.InnerText.Contains("first-page footer"));
        Assert.Contains(main.FooterParts, part => part.Footer!.InnerText.Contains("even-page footer"));
        Assert.NotNull(main.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>());
        Assert.Contains(sectionProperties, section => section.GetFirstChild<TitlePage>() != null);
    }
}
