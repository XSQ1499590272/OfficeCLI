using System.Diagnostics;
using System.Text.Json;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public class WordViewSmokeTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Theory]
    [InlineData("text")]
    [InlineData("annotated")]
    [InlineData("outline")]
    [InlineData("stats")]
    [InlineData("issues")]
    public void ViewJsonModes_ReturnParseableEnvelopesWithStableKeys(string mode)
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "CLI view paragraph" });
        }

        var result = RunOfficeCli("view", path, mode, "--json");

        Assert.True(result.ExitCode == 0, $"exit={result.ExitCode}\nstdout:\n{result.Stdout}\nstderr:\n{result.Stderr}");
        using var doc = JsonDocument.Parse(result.Stdout);
        var root = doc.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), result.Stdout);
        var data = root.GetProperty("data");

        switch (mode)
        {
            case "text":
                Assert.Equal(2, data.GetProperty("totalElements").GetInt32());
                Assert.Contains("CLI view paragraph", data.GetRawText());
                break;
            case "annotated":
                Assert.Equal("annotated", data.GetProperty("view").GetString());
                Assert.Contains("CLI view paragraph", data.GetProperty("content").GetString());
                break;
            case "outline":
                Assert.Equal(1, data.GetProperty("paragraphs").GetInt32());
                Assert.True(data.TryGetProperty("headings", out var headings));
                Assert.Equal(JsonValueKind.Array, headings.ValueKind);
                break;
            case "stats":
                Assert.Equal(1, data.GetProperty("paragraphs").GetInt32());
                Assert.True(data.GetProperty("words").GetInt32() >= 3);
                Assert.True(data.TryGetProperty("styleDistribution", out var styles));
                Assert.Equal(JsonValueKind.Object, styles.ValueKind);
                break;
            case "issues":
                Assert.True(data.TryGetProperty("count", out _));
                Assert.Equal(JsonValueKind.Array, data.GetProperty("issues").ValueKind);
                break;
        }
    }

    [Fact]
    public void ViewJson_CorruptDocxReturnsErrorEnvelopeAndLeavesBytesUntouched()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_corrupt_{Guid.NewGuid():N}.docx");
        var original = new byte[] { 0x6E, 0x6F, 0x74, 0x2D, 0x64, 0x6F, 0x63, 0x78 };
        File.WriteAllBytes(path, original);

        try
        {
            var result = RunOfficeCli("view", path, "text", "--json");

            Assert.Equal(1, result.ExitCode);
            using var doc = JsonDocument.Parse(result.Stdout);
            var root = doc.RootElement;
            Assert.False(root.GetProperty("success").GetBoolean());
            Assert.Equal("corrupt_file", root.GetProperty("error").GetProperty("code").GetString());
            Assert.Equal(original, File.ReadAllBytes(path));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    private static CliRunResult RunOfficeCli(params string[] args)
    {
        var officeCliDll = FindOfficeCliDll();
        var psi = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        psi.ArgumentList.Add(officeCliDll);
        foreach (var arg in args) psi.ArgumentList.Add(arg);
        psi.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";

        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start officecli");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        if (!process.WaitForExit(10_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("officecli process did not exit within 10 seconds");
        }

        return new CliRunResult(process.ExitCode, stdoutTask.GetAwaiter().GetResult(), stderrTask.GetAwaiter().GetResult());
    }

    private static string FindOfficeCliDll()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "officecli.slnx")))
        {
            dir = dir.Parent;
        }

        if (dir == null)
            throw new InvalidOperationException("Could not locate officecli.slnx from test output directory");

        var candidates = Directory.GetFiles(
                Path.Combine(dir.FullName, "src", "officecli", "bin"),
                "officecli.dll",
                SearchOption.AllDirectories)
            .Where(path => File.Exists(Path.Combine(Path.GetDirectoryName(path)!, "officecli.runtimeconfig.json")))
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .ToList();

        return candidates.FirstOrDefault()
            ?? throw new InvalidOperationException("Could not locate built officecli.dll under src/officecli/bin");
    }

    private sealed record CliRunResult(int ExitCode, string Stdout, string Stderr);
}
