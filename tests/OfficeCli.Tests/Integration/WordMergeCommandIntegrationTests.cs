using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordMergeCommandIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void MergeCommand_ReplacesInlineNestedAndArrayPlaceholders_AndKeepsUnresolvedVisible()
    {
        var template = CreateTemplate("Hello {{name}} / {{user.name}} / {{items[0]}} / {{missing}}");
        var output = TempPath("merged");

        try
        {
            var exitCode = Invoke(
                "merge", template, output,
                "--data", "{\"name\":\"Alice\",\"user\":{\"name\":\"Bob\"},\"items\":[\"one\",\"two\"]}");

            Assert.Equal(0, exitCode);
            using var handler = new WordHandler(output, editable: false);
            Assert.Equal(
                "Hello Alice / Bob / one / {{missing}}",
                handler.Query("paragraph").Single().Text);
            Assert.Empty(handler.Validate());
        }
        finally
        {
            TryDelete(output);
        }
    }

    [Fact]
    public void MergeCommand_AcceptsJsonDataFile_AndRefusesExistingOutputWithoutForce()
    {
        var template = CreateTemplate("File value: {{value}}");
        var dataFile = TempPath("merge-data", ".json");
        var output = TempPath("merge-output");
        File.WriteAllText(dataFile, "{\"value\":\"from-file\"}");

        try
        {
            Assert.Equal(0, Invoke("merge", template, output, "--data", dataFile));
            using (var merged = new WordHandler(output, editable: false))
                Assert.Equal("File value: from-file", merged.Query("paragraph").Single().Text);

            var before = File.ReadAllBytes(output);
            Assert.NotEqual(0, Invoke("merge", template, output, "--data", "{\"value\":\"blocked\"}"));
            Assert.Equal(before, File.ReadAllBytes(output));

            Assert.Equal(0, Invoke("merge", template, output, "--data", "{\"value\":\"forced\"}", "--force"));
            using var forced = new WordHandler(output, editable: false);
            Assert.Equal("File value: forced", forced.Query("paragraph").Single().Text);
        }
        finally
        {
            TryDelete(dataFile);
            TryDelete(output);
        }
    }

    [Fact]
    public void MergeCommand_InvalidJsonFailsBeforeCreatingOutput()
    {
        var template = CreateTemplate("{{value}}");
        var output = TempPath("merge-invalid");

        try
        {
            var exitCode = Invoke("merge", template, output, "--data", "[]");

            Assert.NotEqual(0, exitCode);
            Assert.False(File.Exists(output));
        }
        finally
        {
            TryDelete(output);
        }
    }

    private string CreateTemplate(string text)
    {
        var path = CreateBlankDocx();
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = text });
        return path;
    }

    private static int Invoke(params string[] args)
    {
        var root = CommandBuilder.BuildRootCommand();
        return root.Parse(args).Invoke();
    }

    private static string TempPath(string prefix, string extension = ".docx")
        => Path.Combine(Path.GetTempPath(), $"officecli_{prefix}_{Guid.NewGuid():N}{extension}");

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }
}
