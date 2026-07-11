using System.Text.Json;
using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Collection("Word CLI output")]
[Trait("Speed", "Integration")]
public sealed class WordViewCommandIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void ViewJsonModes_ReturnParseableEnvelopesWithRangeOptions()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
        {
            setup.Add("/body", "paragraph", null, new() { ["text"] = "view one" });
            setup.Add("/body", "paragraph", null, new() { ["text"] = "view two" });
        }

        AssertEnvelope(InvokeCaptured("view", path, "text", "--max-lines", "1", "--json"));
        AssertEnvelope(InvokeCaptured("view", path, "annotated", "--start", "1", "--end", "1", "--json"));
        AssertEnvelope(InvokeCaptured("view", path, "outline", "--json"));
        AssertEnvelope(InvokeCaptured("view", path, "stats", "--json"));
        AssertEnvelope(InvokeCaptured("view", path, "issues", "--limit", "1", "--json"));
        AssertEnvelope(InvokeCaptured("view", path, "forms", "--json"));
    }

    [Fact]
    public void ViewHtml_OutFileAndInvalidRenderOptionHaveStableBoundaries()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
            setup.Add("/body", "paragraph", null, new() { ["text"] = "HTML view" });

        var output = Path.Combine(Path.GetTempPath(), $"officecli_view_{Guid.NewGuid():N}.html");
        try
        {
            var html = InvokeCaptured("view", path, "html", "--page", "1", "--out", output);
            Assert.Equal(0, html.ExitCode);
            Assert.True(string.IsNullOrWhiteSpace(html.Stderr), html.Stderr);
            Assert.True(File.Exists(output));
            Assert.Contains("HTML view", File.ReadAllText(output), StringComparison.Ordinal);

            var invalid = InvokeCaptured("view", path, "text", "--render", "invalid");
            Assert.NotEqual(0, invalid.ExitCode);
        }
        finally
        {
            try { if (File.Exists(output)) File.Delete(output); } catch { }
        }
    }

    private static void AssertEnvelope(CommandResult result)
    {
        Assert.Equal(0, result.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(result.Stderr), result.Stderr);
        using var json = JsonDocument.Parse(result.Stdout);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
    }

    private static CommandResult InvokeCaptured(params string[] args)
    {
        lock (typeof(WordViewCommandIntegrationTests))
        {
            var oldOut = Console.Out;
            var oldError = Console.Error;
            using var stdout = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
            using var stderr = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
            Console.SetOut(stdout);
            Console.SetError(stderr);
            try
            {
                var root = CommandBuilder.BuildRootCommand();
                var exitCode = root.Parse(args).Invoke();
                return new CommandResult(exitCode, stdout.ToString(), stderr.ToString());
            }
            finally
            {
                Console.SetOut(oldOut);
                Console.SetError(oldError);
            }
        }
    }

    private sealed record CommandResult(int ExitCode, string Stdout, string Stderr);
}
