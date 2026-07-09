using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordDocumentSettingsTests : WordTestBase
{
    [Fact]
    public void SetDocumentPropertiesDefaultsAndSettings_ReadsBackCanonicalKeys()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/", new()
        {
            ["creator"] = "Office Agent",
            ["title"] = "Unit Title",
            ["docDefaults.font"] = "Georgia",
            ["docDefaults.fontSize"] = "13pt",
            ["docGrid.type"] = "linesAndChars",
            ["docGrid.linePitch"] = "360",
            ["charSpacingControl"] = "doNotCompress"
        });

        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "inherits" });
        var document = handler.Get("/");
        var paragraph = handler.Get(paragraphPath);

        Assert.Equal("Office Agent", Fmt(document)["author"]);
        Assert.Equal("Unit Title", Fmt(document)["title"]);
        Assert.Equal("linesAndChars", Fmt(document)["docGrid.type"]);
        Assert.Equal(360, Fmt(document)["docGrid.linePitch"]);
        Assert.Equal("doNotCompress", Fmt(document)["charSpacingControl"]);
        Assert.Equal("13pt", Fmt(paragraph)["effective.size"]);
        Assert.Equal("Georgia", Fmt(paragraph)["effective.font.ascii"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void DocDefaultsAndSettingsPaths_ReadWriteTheirOwnNodes()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/docDefaults", new()
        {
            ["font"] = "Aptos",
            ["font.eastAsia"] = "Songti SC",
            ["fontSize"] = "12pt",
            ["bold"] = "true"
        });
        handler.Set("/settings", new()
        {
            ["docGrid.type"] = "lines",
            ["docGrid.linePitch"] = "240"
        });

        var docDefaults = handler.Get("/docDefaults");
        var document = handler.Get("/");

        Assert.Equal("docDefaults", docDefaults.Type);
        Assert.Equal("Aptos", Fmt(docDefaults)["docDefaults.font"]);
        Assert.Equal("Songti SC", Fmt(docDefaults)["docDefaults.font.eastAsia"]);
        Assert.Equal("12pt", Fmt(docDefaults)["docDefaults.fontSize"]);
        Assert.Equal(true, Fmt(docDefaults)["docDefaults.bold"]);
        Assert.Equal("lines", Fmt(document)["docGrid.type"]);
        Assert.Equal(240, Fmt(document)["docGrid.linePitch"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void DocDefaultsParagraphProperties_ReadBackAndAffectParagraph()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/docDefaults", new()
        {
            ["rtl"] = "true",
            ["alignment"] = "right",
            ["spaceBefore"] = "6pt",
            ["spaceAfter"] = "12pt",
            ["lineSpacing"] = "1.5"
        });

        var docDefaults = handler.Get("/docDefaults");

        Assert.Equal(true, Fmt(docDefaults)["docDefaults.rtl"]);
        Assert.Equal("right", Fmt(docDefaults)["docDefaults.alignment"]);
        Assert.Equal("6pt", Fmt(docDefaults)["docDefaults.spaceBefore"]);
        Assert.Equal("12pt", Fmt(docDefaults)["docDefaults.spaceAfter"]);
        Assert.Equal("1.5x", Fmt(docDefaults)["docDefaults.lineSpacing"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void CompatibilitySettings_ReadBackModeFlagAndPreset()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/settings", new()
        {
            ["compatibility.mode"] = "14",
            ["compatibility.usePrinterMetrics"] = "true"
        });
        var document = handler.Get("/");

        Assert.Equal(14, Fmt(document)["compatibility.mode"]);
        Assert.Equal(true, Fmt(document)["compatibility.usePrinterMetrics"]);

        handler.Set("/settings", new() { ["compatibility.preset"] = "word2019" });
        var presetDocument = handler.Get("/");

        Assert.Equal(15, Fmt(presetDocument)["compatibility.mode"]);
        Assert.Equal(true, Fmt(presetDocument)["compatibility.useFarEastLayout"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddSection_ReadsBackLayoutMarginsAndColumns()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var sectionPath = handler.Add("/body", "section", null, new()
        {
            ["type"] = "continuous",
            ["orientation"] = "landscape",
            ["marginTop"] = "1cm",
            ["marginBottom"] = "1cm",
            ["columns"] = "2",
            ["columnSpace"] = "1cm"
        });

        var section = handler.Get(sectionPath);

        Assert.Equal("section", section.Type);
        Assert.Equal("continuous", Fmt(section)["type"]);
        Assert.Equal("landscape", Fmt(section)["orientation"]);
        Assert.Equal("1cm", Fmt(section)["marginTop"]);
        Assert.Equal("1cm", Fmt(section)["marginBottom"]);
        Assert.Equal(2, Convert.ToInt32(Fmt(section)["columns"]));
        Assert.Equal("1cm", Fmt(section)["columnSpace"]);
    }

    [Fact]
    public void AddSection_ReadsBackPageSizeDirectionAndRtlGutter()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var sectionPath = handler.Add("/body", "section", null, new()
        {
            ["pageWidth"] = "20cm",
            ["pageHeight"] = "10cm",
            ["direction"] = "rtl",
            ["rtlGutter"] = "true"
        });

        var section = handler.Get(sectionPath);

        Assert.Equal("20cm", Fmt(section)["pageWidth"]);
        Assert.Equal("10cm", Fmt(section)["pageHeight"]);
        Assert.Equal("rtl", Fmt(section)["direction"]);
        Assert.Equal(true, Fmt(section)["rtlGutter"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddSection_NormalizesLengthUnitsToCentimeterReadBack()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var sectionPath = handler.Add("/body", "section", null, new()
        {
            ["pageWidth"] = "8.5in",
            ["pageHeight"] = "792pt",
            ["marginTop"] = "1440dxa",
            ["marginBottom"] = "72pt",
            ["columnSpace"] = "0.5in",
            ["columns"] = "2"
        });

        var section = handler.Get(sectionPath);

        Assert.Equal("21.59cm", Fmt(section)["pageWidth"]);
        Assert.Equal("27.94cm", Fmt(section)["pageHeight"]);
        Assert.Equal("2.54cm", Fmt(section)["marginTop"]);
        Assert.Equal("2.54cm", Fmt(section)["marginBottom"]);
        Assert.Equal("1.27cm", Fmt(section)["columnSpace"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void RtlLocaleBlank_ReadsDocumentLanguageAndParagraphInheritedDirection()
    {
        var path = CreateBlankDocx("ar-SA");

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "مرحبا" });

        var document = handler.Get("/");
        var paragraph = handler.Get(paragraphPath);

        Assert.Equal("ar-SA", Fmt(document)["lang.cs"]);
        Assert.Equal("ar-SA", Fmt(document)["locale"]);
        Assert.Equal("rtl", Fmt(paragraph)["effective.direction"]);
        Assert.Equal(true, Fmt(paragraph)["effective.rtl"]);
        Assert.Equal("/section[1]", Fmt(paragraph)["effective.direction.src"]);
        Assert.Empty(handler.Validate());
    }
}
