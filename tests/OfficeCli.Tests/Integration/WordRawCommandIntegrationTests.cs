using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordRawCommandIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddPartCommand_CreatesChartPartAndKeepsPackageValid()
    {
        var path = CreateBlankDocx();

        var exitCode = Invoke(
            "add-part", path, "/", "--type", "chart");

        Assert.Equal(0, exitCode);
        using var handler = new WordHandler(path, editable: false);
        Assert.Contains("<c:chartSpace", handler.Raw("/chart[1]"), StringComparison.Ordinal);
        var validationError = Assert.Single(handler.Validate());
        Assert.Equal("Schema", validationError.ErrorType);
        Assert.Contains("incomplete content", validationError.Description, StringComparison.Ordinal);
    }

    [Fact]
    public void RawSetCommand_WritesNamespacedAttributeAndPersistsIt()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
            setup.Add("/body", "paragraph", null, new() { ["text"] = "raw command" });

        var exitCode = Invoke(
            "raw-set", path, "/document",
            "--xpath", "//w:body/w:p[1]",
            "--action", "setattr",
            "--xml", "w:rsidR=00112233");

        Assert.Equal(0, exitCode);
        using var reopened = new WordHandler(path, editable: false);
        Assert.Contains("00112233", reopened.Raw("/document"), StringComparison.Ordinal);
        Assert.Empty(reopened.Validate());
    }

    [Fact]
    public void RawSetCommand_InvalidXPathLeavesPackageBytesUnchanged()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
            setup.Add("/body", "paragraph", null, new() { ["text"] = "unchanged" });
        string beforeXml;
        using (var beforeHandler = new WordHandler(path, editable: false))
            beforeXml = beforeHandler.Raw("/document");

        var exitCode = Invoke(
            "raw-set", path, "/document",
            "--xpath", "//w:body/w:p[99]",
            "--action", "setattr",
            "--xml", "w:rsidR=44556677");

        Assert.NotEqual(0, exitCode);
        using var reopened = new WordHandler(path, editable: false);
        Assert.Equal(beforeXml, reopened.Raw("/document"));
        Assert.Equal("unchanged", reopened.Query("paragraph").Single().Text);
    }

    private static int Invoke(params string[] args)
    {
        var root = CommandBuilder.BuildRootCommand();
        return root.Parse(args).Invoke();
    }
}
