using System.Text.Json;
using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Collection("Word CLI output")]
[Trait("Speed", "Integration")]
public sealed class WordMutationCommandIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddCommand_SupportsAppendIndexBeforeAfter_AndRejectsConflictingPositions()
    {
        var path = CreateBlankDocx();

        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--prop", "text=first"));
        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--prop", "text=last"));
        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--index", "1", "--prop", "text=index"));
        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--before", "/body/p[1]", "--prop", "text=before"));
        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--after", "/body/p[4]", "--prop", "text=after"));

        using (var handler = new WordHandler(path, editable: false))
        {
            Assert.Equal(
                new[] { "before", "first", "index", "last", "after" },
                handler.Query("paragraph").Select(node => node.Text));
        }

        var beforeInvalid = File.ReadAllBytes(path);
        var invalidExit = Invoke(
            "add", path, "/body", "--type", "paragraph",
            "--index", "0", "--before", "/body/p[1]", "--prop", "text=invalid");

        Assert.NotEqual(0, invalidExit);
        Assert.Equal(beforeInvalid, File.ReadAllBytes(path));
    }

    [Fact]
    public void SetCommand_AppliesCaseInsensitiveLastValueAlongsideSupportedAndUnsupportedProps()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
            setup.Add("/body", "paragraph", null, new() { ["text"] = "original" });

        var exitCode = Invoke(
            "set", path, "/body/p[1]",
            "--prop", "TEXT=first",
            "--prop", "text=second",
            "--prop", "ALIGN=center",
            "--prop", "unsupportedProperty=value");

        Assert.Equal(2, exitCode);
        using (var handler = new WordHandler(path, editable: false))
        {
            var paragraph = handler.Get("/body/p[1]");
            Assert.Equal("second", paragraph.Text);
            Assert.Equal("center", paragraph.Format["align"]);
        }

        var beforeMissingProps = File.ReadAllBytes(path);
        Assert.NotEqual(0, Invoke("set", path, "/body/p[1]"));
        Assert.Equal(beforeMissingProps, File.ReadAllBytes(path));
    }

    [Fact]
    public void MutationArgumentFailures_DoNotChangeTheDocument()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
            setup.Add("/body", "paragraph", null, new() { ["text"] = "source" });

        var before = File.ReadAllBytes(path);

        Assert.NotEqual(0, Invoke("add", path, "/body"));
        Assert.Equal(before, File.ReadAllBytes(path));

        Assert.NotEqual(0, Invoke("add", path, "/body", "--from", "/body/p[1]", "--prop", "text=ignored"));
        Assert.Equal(before, File.ReadAllBytes(path));

        Assert.NotEqual(0, Invoke("add", path, "/body", "--type", "paragraph", "--prop", "=bad"));
        Assert.Equal(before, File.ReadAllBytes(path));

        Assert.NotEqual(0, Invoke("set", path, "paragraph", "--prop", "text=unsafe"));
        Assert.Equal(before, File.ReadAllBytes(path));

        Assert.NotEqual(0, Invoke("remove", path, "paragraph"));
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Fact]
    public void GetAndQueryJson_ReturnStructuredResults_AndSaveRejectsTextNodes()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
        {
            setup.Add("/body", "paragraph", null, new() { ["text"] = "alpha" });
            setup.Add("/body", "paragraph", null, new() { ["text"] = "beta" });
        }

        var get = InvokeCaptured("get", path, "/body/p[1]", "--depth", "0", "--json");
        Assert.Equal(0, get.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(get.Stderr), get.Stderr);
        AssertResultsArray(get.Stdout, expectedCount: 1);

        var query = InvokeCaptured("query", path, "paragraph", "--find", "alpha", "--json");
        Assert.Equal(0, query.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(query.Stderr), query.Stderr);
        AssertResultsArray(query.Stdout, expectedCount: 1);

        var outputPath = Path.Combine(Path.GetTempPath(), $"officecli_word_text_{Guid.NewGuid():N}.bin");
        try
        {
            var save = InvokeCaptured("get", path, "/body/p[1]", "--save", outputPath, "--json");
            Assert.NotEqual(0, save.ExitCode);
            Assert.False(File.Exists(outputPath));
        }
        finally
        {
            try { if (File.Exists(outputPath)) File.Delete(outputPath); } catch { }
        }
    }

    [Fact]
    public void RemoveMoveCopyAndSwapCommands_PreserveExpectedBodyOrder()
    {
        var path = CreateBlankDocx();
        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--prop", "text=A"));
        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--prop", "text=B"));
        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--prop", "text=C"));

        Assert.Equal(0, Invoke("move", path, "/body/p[3]", "--before", "/body/p[1]"));
        AssertBodyText(path, "C", "A", "B");

        Assert.Equal(0, Invoke("add", path, "/body", "--from", "/body/p[2]", "--after", "/body/p[3]"));
        AssertBodyText(path, "C", "A", "B", "A");

        Assert.Equal(0, Invoke("swap", path, "/body/p[1]", "/body/p[4]"));
        AssertBodyText(path, "A", "A", "B", "C");

        Assert.Equal(0, Invoke("remove", path, "/body/p[2]"));
        AssertBodyText(path, "A", "B", "C");

        var beforeInvalid = ReadDocumentXml(path);
        Assert.NotEqual(0, Invoke("move", path, "/body/p[99]", "--before", "/body/p[1]"));
        Assert.Equal(beforeInvalid, ReadDocumentXml(path));
        Assert.NotEqual(0, Invoke("swap", path, "/body/p[99]", "/body/p[1]"));
        Assert.Equal(beforeInvalid, ReadDocumentXml(path));
        Assert.NotEqual(0, Invoke("remove", path, "/body/p[99]"));
        Assert.Equal(beforeInvalid, ReadDocumentXml(path));

        using var final = new WordHandler(path, editable: false);
        Assert.Empty(final.Validate());
    }

    [Fact]
    public void SaveCloseAndValidate_AreIdempotentWithoutAResident()
    {
        var path = CreateBlankDocx();
        var before = File.ReadAllBytes(path);

        Assert.Equal(0, Invoke("save", path));
        Assert.Equal(0, Invoke("save", path));
        Assert.Equal(0, Invoke("close", path));
        Assert.Equal(0, Invoke("close", path));
        Assert.Equal(0, Invoke("validate", path));
        Assert.Equal(0, Invoke("validate", path));

        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Fact]
    public void JsonWordCommands_KeepStdoutParseableAndDiagnosticsOffStdout()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
        {
            setup.Add("/body", "paragraph", null, new() { ["text"] = "json-one" });
            setup.Add("/body", "paragraph", null, new() { ["text"] = "json-two" });
            setup.Add("/body", "paragraph", null, new() { ["text"] = "json-three" });
        }

        AssertJsonSuccess(InvokeCaptured("get", path, "/body/p[1]", "--json"));
        AssertJsonSuccess(InvokeCaptured("query", path, "paragraph", "--json"));
        AssertJsonSuccess(InvokeCaptured("set", path, "/body/p[1]", "--prop", "text=updated", "--json"));
        AssertJsonSuccess(InvokeCaptured("add", path, "/body", "--type", "paragraph", "--prop", "text=json-four", "--json"));
        AssertJsonSuccess(InvokeCaptured("move", path, "/body/p[4]", "--before", "/body/p[1]", "--json"));
        AssertJsonSuccess(InvokeCaptured("swap", path, "/body/p[1]", "/body/p[2]", "--json"));
        AssertJsonSuccess(InvokeCaptured("remove", path, "/body/p[1]", "--json"));
        AssertJsonSuccess(InvokeCaptured("validate", path, "--json"));
        AssertJsonSuccess(InvokeCaptured("save", path, "--json"));
        AssertJsonSuccess(InvokeCaptured("close", path, "--json"));

        var unsupported = InvokeCaptured(
            "set", path, "/body/p[1]", "--prop", "unsupportedProperty=value", "--json");
        Assert.Equal(2, unsupported.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(unsupported.Stderr), unsupported.Stderr);
        using var errorJson = JsonDocument.Parse(unsupported.Stdout);
        Assert.False(errorJson.RootElement.GetProperty("success").GetBoolean());
    }

    private static int Invoke(params string[] args)
    {
        var root = CommandBuilder.BuildRootCommand();
        return root.Parse(args).Invoke();
    }

    private static CommandResult InvokeCaptured(params string[] args)
    {
        lock (ConsoleGate.Instance)
        {
            var oldOut = Console.Out;
            var oldError = Console.Error;
            using var stdout = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
            using var stderr = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
            Console.SetOut(stdout);
            Console.SetError(stderr);
            try
            {
                return new CommandResult(Invoke(args), stdout.ToString(), stderr.ToString());
            }
            finally
            {
                Console.SetOut(oldOut);
                Console.SetError(oldError);
            }
        }
    }

    private static void AssertResultsArray(string output, int expectedCount)
    {
        using var json = JsonDocument.Parse(output);
        var root = json.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), output);
        var results = root.GetProperty("data").GetProperty("results");
        Assert.Equal(JsonValueKind.Array, results.ValueKind);
        Assert.Equal(expectedCount, results.GetArrayLength());
    }

    private static void AssertJsonSuccess(CommandResult result)
    {
        Assert.Equal(0, result.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(result.Stderr), result.Stderr);
        using var json = JsonDocument.Parse(result.Stdout);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
    }

    private static void AssertBodyText(string path, params string[] expected)
    {
        using var handler = new WordHandler(path, editable: false);
        Assert.Equal(expected, handler.Query("paragraph").Select(node => node.Text));
    }

    private static string ReadDocumentXml(string path)
    {
        using var handler = new WordHandler(path, editable: false);
        return handler.Raw("/document");
    }

    private sealed record CommandResult(int ExitCode, string Stdout, string Stderr);

    private static class ConsoleGate
    {
        public static readonly object Instance = new();
    }
}

[CollectionDefinition("Word CLI output", DisableParallelization = true)]
public sealed class WordCliOutputCollection
{
}
