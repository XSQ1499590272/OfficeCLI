using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordValidateCommandIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void ValidateCommand_HealthyDocumentReturnsSuccess()
    {
        var path = CreateBlankDocx();

        var exitCode = Invoke("validate", path, "--json");

        Assert.Equal(0, exitCode);
    }

    [Fact]
    public void ValidateCommand_SchemaErrorReturnsFailure()
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "invalid" });
            handler.RawSet(
                "/document",
                "//w:body/w:p[1]",
                "append",
                "<w:notAWordChild xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" />");
        }

        var exitCode = Invoke("validate", path, "--json");

        Assert.NotEqual(0, exitCode);
    }

    private static int Invoke(params string[] args)
    {
        var root = CommandBuilder.BuildRootCommand();
        return root.Parse(args).Invoke();
    }
}
