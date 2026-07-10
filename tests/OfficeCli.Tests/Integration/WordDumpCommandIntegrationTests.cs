using System.Text.Json;
using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordDumpCommandIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void DumpCommand_OutFileIsBareDeterministicReplayArray()
    {
        var sourcePath = CreateBlankDocx();
        var firstOut = Path.Combine(Path.GetTempPath(), $"officecli_dump_{Guid.NewGuid():N}_1.json");
        var secondOut = Path.Combine(Path.GetTempPath(), $"officecli_dump_{Guid.NewGuid():N}_2.json");

        try
        {
            using (var handler = new WordHandler(sourcePath, editable: true))
            {
                handler.Add("/body", "paragraph", null, new() { ["text"] = "dump command" });
                handler.Add("/body", "table", null, new() { ["data"] = "A,B;C,D" });
            }

            var firstExit = InvokeDump(sourcePath, "/", firstOut);
            var secondExit = InvokeDump(sourcePath, "/", secondOut);

            Assert.Equal(0, firstExit);
            Assert.Equal(0, secondExit);
            var first = File.ReadAllText(firstOut);
            var second = File.ReadAllText(secondOut);

            Assert.Equal(first, second);
            Assert.EndsWith(Environment.NewLine, first, StringComparison.Ordinal);
            using var json = JsonDocument.Parse(first);
            Assert.Equal(JsonValueKind.Array, json.RootElement.ValueKind);
            Assert.Contains(json.RootElement.EnumerateArray(), item =>
                item.TryGetProperty("command", out var command)
                && item.TryGetProperty("type", out var type)
                && command.GetString() == "add"
                && type.GetString() == "p");
        }
        finally
        {
            TryDelete(firstOut);
            TryDelete(secondOut);
        }
    }

    [Fact]
    public void DumpCommand_InvalidSubtreeDoesNotCreateOutputFile()
    {
        var sourcePath = CreateBlankDocx();
        var outputPath = Path.Combine(Path.GetTempPath(), $"officecli_dump_invalid_{Guid.NewGuid():N}.json");

        try
        {
            var exitCode = InvokeDump(sourcePath, "/body/p[99]", outputPath);

            Assert.NotEqual(0, exitCode);
            Assert.False(File.Exists(outputPath));
        }
        finally
        {
            TryDelete(outputPath);
        }
    }

    private static int InvokeDump(string sourcePath, string subtree, string outputPath)
    {
        var root = CommandBuilder.BuildRootCommand();
        return root.Parse(["dump", sourcePath, subtree, "--out", outputPath]).Invoke();
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }
}
