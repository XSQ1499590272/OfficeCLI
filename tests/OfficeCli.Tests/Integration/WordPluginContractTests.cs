using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Core.Plugins;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordPluginContractTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    [Fact]
    public void DumpReader_StreamsJsonlIntoWordAndRejectsTopLevelArray()
    {
        if (OperatingSystem.IsWindows()) return;

        var pluginPath = CreateTemp("officecli_test_dump_reader", ".sh");
        File.WriteAllText(pluginPath, """
            #!/bin/sh
            if [ "$1" = "--info" ]; then
              printf '%s\n' '{"name":"test-dump-reader","version":"1.0.0","protocol":1,"kinds":["dump-reader"],"extensions":[".doc"],"target":"docx","runtime":"other","idle_timeout_seconds":{"default":5,"verbs":{"dump":5}}}'
              exit 0
            fi
            if [ "$1" = "dump" ]; then
              if grep -q 'bad-array' "$2"; then
                printf '%s\n' '[{"command":"add","parent":"/body","type":"paragraph","props":{"text":"bad"}}]'
              else
                printf '%s\n' '{"command":"add","parent":"/body","type":"paragraph","props":{"text":"plugin paragraph"}}'
              fi
              exit 0
            fi
            exit 1
            """, new UTF8Encoding(false));
        File.SetUnixFileMode(pluginPath,
            UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);

        var sourcePath = CreateTemp("officecli_plugin_source", ".doc");
        File.WriteAllText(sourcePath, "good source");
        var previousPlugin = Environment.GetEnvironmentVariable("OFFICECLI_PLUGIN_DUMP_READER_DOC");
        Environment.SetEnvironmentVariable("OFFICECLI_PLUGIN_DUMP_READER_DOC", pluginPath);
        PluginRegistry.InvalidateCache();

        try
        {
            var converted = DumpReaderInvoker.Run(sourcePath, ".doc");
            Assert.Equal("test-dump-reader", converted.Plugin.Manifest.Name);
            Assert.Equal("docx", converted.Plugin.Manifest.ResolveTargetFormat());
            try
            {
                using var document = new WordHandler(converted.ConvertedPath, editable: false);
                var paragraph = Assert.Single(document.Query("paragraph"));
                Assert.Equal("plugin paragraph", paragraph.Text);
                Assert.Empty(document.Validate());
            }
            finally
            {
                try { File.Delete(converted.ConvertedPath); } catch { }
            }

            var lint = RunOfficeCli("plugins", "lint", pluginPath, "--fixture", sourcePath, "--json");
            Assert.Equal(0, lint.ExitCode);
            using (var lintDocument = JsonDocument.Parse(lint.Stdout))
            {
                var lintRoot = lintDocument.RootElement;
                Assert.True(lintRoot.GetProperty("success").GetBoolean(), lint.Stdout);
                Assert.Equal(0, lintRoot.GetProperty("data").GetProperty("unknown_prop_count").GetInt32());
            }

            File.WriteAllText(sourcePath, "bad-array source");
            var error = Assert.Throws<CliException>(() => DumpReaderInvoker.Run(sourcePath, ".doc"));
            Assert.Equal("corrupt_batch", error.Code);
            Assert.Contains("JSON array", error.Message);

            var badLint = RunOfficeCli("plugins", "lint", pluginPath, "--fixture", sourcePath, "--json");
            Assert.Equal(1, badLint.ExitCode);
            using (var badLintDocument = JsonDocument.Parse(badLint.Stdout))
            {
                Assert.False(badLintDocument.RootElement.GetProperty("success").GetBoolean());
                Assert.Equal("corrupt_batch", badLintDocument.RootElement.GetProperty("error").GetProperty("code").GetString());
            }
        }
        finally
        {
            Environment.SetEnvironmentVariable("OFFICECLI_PLUGIN_DUMP_READER_DOC", previousPlugin);
            PluginRegistry.InvalidateCache();
        }
    }

    [Fact]
    public void PluginManifest_InvalidKindTargetAndTimeoutExposeWarningsAndSafeFallback()
    {
        var manifest = new PluginManifest
        {
            Name = "invalid-word-plugin",
            Version = "0.0.1",
            Protocol = 1,
            Kinds = ["dump-reader", "unknown-kind"],
            Extensions = [".doc"],
            Target = "pdf",
            IdleTimeoutSeconds = new PluginIdleTimeout { Default = 0 }
        };

        var warnings = manifest.Warnings();

        Assert.Contains(warnings, warning => warning.Contains("unknown kind", StringComparison.Ordinal));
        Assert.Contains(warnings, warning => warning.Contains("target", StringComparison.Ordinal));
        Assert.Contains(warnings, warning => warning.Contains("idle_timeout_seconds.default", StringComparison.Ordinal));
        Assert.Equal(60, manifest.ResolveIdleTimeout("dump"));
    }

    [Fact]
    public void WordCommand_WithInstallAndUpdateGuardsKeepsProtocolStreamsClean()
    {
        var documentPath = CreateTemp("officecli_guarded_word", ".docx");
        OfficeCli.BlankDocCreator.Create(documentPath);

        var result = RunOfficeCli("validate", documentPath, "--json");

        Assert.Equal(0, result.ExitCode);
        using var json = JsonDocument.Parse(result.Stdout);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
        Assert.DoesNotContain("install", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("update", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }

    private string CreateTemp(string prefix, string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{prefix}_{Guid.NewGuid():N}{extension}");
        _tempFiles.Add(path);
        return path;
    }

    private static CliRunResult RunOfficeCli(params string[] args)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "officecli.slnx")))
            dir = dir.Parent;
        Assert.NotNull(dir);

        var dll = Directory.GetFiles(
                Path.Combine(dir!.FullName, "src", "officecli", "bin"),
                "officecli.dll", SearchOption.AllDirectories)
            .Where(path => File.Exists(Path.Combine(Path.GetDirectoryName(path)!, "officecli.runtimeconfig.json")))
            .OrderByDescending(IsCurrentHostRuntimePath)
            .ThenByDescending(File.GetLastWriteTimeUtc)
            .FirstOrDefault();
        Assert.NotNull(dll);

        var psi = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        psi.ArgumentList.Add(dll!);
        foreach (var arg in args) psi.ArgumentList.Add(arg);
        psi.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";
        psi.Environment["OFFICECLI_SKIP_UPDATE"] = "1";
        psi.Environment["OFFICECLI_NO_AUTO_INSTALL"] = "1";

        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start officecli");
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit(30_000);
        return new CliRunResult(process.ExitCode, stdout, stderr);
    }

    private static bool IsCurrentHostRuntimePath(string path)
    {
        var normalized = path.Replace('\\', '/');
        var marker = OperatingSystem.IsWindows()
            ? "/win-"
            : OperatingSystem.IsMacOS()
                ? "/osx-"
                : "/linux-";
        return normalized.Contains(marker, StringComparison.OrdinalIgnoreCase);
    }

    private sealed record CliRunResult(int ExitCode, string Stdout, string Stderr);

    public void Dispose()
    {
        foreach (var path in _tempFiles)
        {
            try { File.Delete(path); } catch { }
        }
    }
}
