using System.Diagnostics;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public class WordExampleScriptSmokeTests
{
    [Fact]
    public void WordShellExamples_GenerateInspectableDocxInTemporaryDirectory()
    {
        if (OperatingSystem.IsWindows())
            return;

        var repoRoot = FindRepositoryRoot();
        var sourceDirectory = Path.Combine(repoRoot, "examples", "word");
        var tempDirectory = Path.Combine(Path.GetTempPath(), $"officecli_word_examples_{Guid.NewGuid():N}");
        CopyDirectory(sourceDirectory, tempDirectory);

        try
        {
            var appHost = FindOfficeCliAppHost(repoRoot);
            var scripts = new[]
            {
                "content-controls.sh",
                "pictures.sh",
                "revisions.sh",
                "tables.sh",
            };

            foreach (var scriptName in scripts)
            {
                var script = Path.Combine(tempDirectory, scriptName);
                Assert.True(File.Exists(script), $"Missing example script {scriptName}.");
                var result = RunScript(script, Path.GetDirectoryName(appHost)!);
                Assert.True(result.ExitCode is 0 or 2,
                    $"{scriptName} failed with exit {result.ExitCode}.\n" +
                    $"stdout:\n{result.Stdout}\nstderr:\n{result.Stderr}");

                var output = Path.ChangeExtension(script, ".docx");
                Assert.True(File.Exists(output),
                    $"{scriptName} did not create {Path.GetFileName(output)}.");
                Assert.True(new FileInfo(output).Length > 0,
                    $"{Path.GetFileName(output)} is empty.");

                using var handler = new WordHandler(output, editable: false);
                Assert.NotEmpty(handler.Query("paragraph"));
            }
        }
        finally
        {
            try { Directory.Delete(tempDirectory, recursive: true); } catch { }
        }
    }

    private static ScriptResult RunScript(string script, string appHostDirectory)
    {
        var startInfo = new ProcessStartInfo("bash")
        {
            WorkingDirectory = Path.GetDirectoryName(script)!,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        startInfo.ArgumentList.Add(script);
        startInfo.Environment["PATH"] = appHostDirectory + Path.PathSeparator +
            (Environment.GetEnvironmentVariable("PATH") ?? string.Empty);
        startInfo.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";
        startInfo.Environment["OFFICECLI_SKIP_UPDATE"] = "1";

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"Could not start {script}");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        if (!process.WaitForExit(180_000))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException($"{Path.GetFileName(script)} did not finish within 180 seconds.");
        }

        return new ScriptResult(
            process.ExitCode,
            stdoutTask.GetAwaiter().GetResult(),
            stderrTask.GetAwaiter().GetResult());
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "officecli.slnx")))
            directory = directory.Parent;

        return directory?.FullName
            ?? throw new InvalidOperationException("Could not locate officecli.slnx from test output directory.");
    }

    private static string FindOfficeCliAppHost(string repoRoot)
    {
        var executableName = OperatingSystem.IsWindows() ? "officecli.exe" : "officecli";
        return Directory.GetFiles(
                Path.Combine(repoRoot, "src", "officecli", "bin"),
                executableName,
                SearchOption.AllDirectories)
            .Where(path => new FileInfo(path).Length > 0)
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .FirstOrDefault()
            ?? throw new InvalidOperationException("Could not locate built officecli apphost.");
    }

    private static void CopyDirectory(string source, string destination)
    {
        Directory.CreateDirectory(destination);
        foreach (var directory in Directory.GetDirectories(source, "*", SearchOption.AllDirectories))
            Directory.CreateDirectory(Path.Combine(destination, Path.GetRelativePath(source, directory)));

        foreach (var file in Directory.GetFiles(source, "*", SearchOption.AllDirectories))
        {
            var target = Path.Combine(destination, Path.GetRelativePath(source, file));
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Copy(file, target, overwrite: true);
        }
    }

    private sealed record ScriptResult(int ExitCode, string Stdout, string Stderr);
}
