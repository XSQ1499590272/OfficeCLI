using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordDocumentContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void SetDocumentProperties_ReadsBackCanonicalKeys()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/", new()
        {
            ["creator"] = "Office Agent",
            ["title"] = "Contract Title",
            ["keywords"] = "office,test",
            ["description"] = "Contract description",
            ["lastModifiedBy"] = "Codex"
        });

        var document = handler.Get("/");

        Assert.Equal("Office Agent", Fmt(document)["author"]);
        Assert.Equal("Contract Title", Fmt(document)["title"]);
        Assert.Equal("office,test", Fmt(document)["keywords"]);
        Assert.Equal("Contract description", Fmt(document)["description"]);
        Assert.Equal("Codex", Fmt(document)["lastModifiedBy"]);
    }

    [Fact]
    public void SetDocDefaults_AffectsNewParagraphEffectiveRunFormatting()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/", new()
        {
            ["docDefaults.font"] = "Georgia",
            ["docDefaults.font.eastAsia"] = "Songti SC",
            ["docDefaults.fontSize"] = "13pt"
        });

        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "inherits defaults" });
        var paragraph = handler.Get(paragraphPath);

        Assert.Equal("13pt", Fmt(paragraph)["effective.size"]);
        Assert.Equal("/docDefaults", Fmt(paragraph)["effective.size.src"]);
        Assert.Equal("Georgia", Fmt(paragraph)["effective.font.ascii"]);
        Assert.Equal("/docDefaults", Fmt(paragraph)["effective.font.ascii.src"]);
        Assert.Equal("Songti SC", Fmt(paragraph)["effective.font.eastAsia"]);
        Assert.Equal("/docDefaults", Fmt(paragraph)["effective.font.eastAsia.src"]);
    }

    [Fact]
    public void SetDocumentGridAndCharacterSpacing_ReadsBackRootSettings()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/", new()
        {
            ["docGrid.type"] = "linesAndChars",
            ["docGrid.linePitch"] = "360",
            ["docGrid.charSpace"] = "120",
            ["charSpacingControl"] = "doNotCompress"
        });

        var document = handler.Get("/");

        Assert.Equal("linesAndChars", Fmt(document)["docGrid.type"]);
        Assert.Equal(360, Fmt(document)["docGrid.linePitch"]);
        Assert.Equal(120, Fmt(document)["docGrid.charSpace"]);
        Assert.Equal("doNotCompress", Fmt(document)["charSpacingControl"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetDocDefaultsPath_WithBareKeysReadsBackDocDefaultsNode()
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

        var docDefaults = handler.Get("/docDefaults");

        Assert.Equal("docDefaults", docDefaults.Type);
        Assert.Equal("Aptos", Fmt(docDefaults)["docDefaults.font"]);
        Assert.Equal("Songti SC", Fmt(docDefaults)["docDefaults.font.eastAsia"]);
        Assert.Equal("12pt", Fmt(docDefaults)["docDefaults.fontSize"]);
        Assert.Equal(true, Fmt(docDefaults)["docDefaults.bold"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetSettingsPath_ReadsBackSettingsFromRoot()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/settings", new()
        {
            ["docGrid.type"] = "lines",
            ["docGrid.linePitch"] = "240",
            ["charSpacingControl"] = "compressPunctuation"
        });

        var document = handler.Get("/");

        Assert.Equal("lines", Fmt(document)["docGrid.type"]);
        Assert.Equal(240, Fmt(document)["docGrid.linePitch"]);
        Assert.Equal("compressPunctuation", Fmt(document)["charSpacingControl"]);
        Assert.Empty(handler.Validate());
    }
}
