using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public class WordPublishedBinarySmokeTests
{
    [Fact]
    public void PublishedMacAndWindowsBinaries_ExposeWordSmokeSurface()
    {
        var repoRoot = FindRepositoryRoot();
        var publishRoot = Path.Combine(Path.GetTempPath(), $"officecli_publish_smoke_{Guid.NewGuid():N}");
        Directory.CreateDirectory(publishRoot);

        try
        {
            var macBinary = Publish(repoRoot, "osx-arm64", Path.Combine(publishRoot, "osx"));
            var windowsBinary = Publish(repoRoot, "win-x64", Path.Combine(publishRoot, "win"));

            Assert.True(new FileInfo(macBinary).Length > 1_000_000, macBinary);
            Assert.True(new FileInfo(windowsBinary).Length > 1_000_000, windowsBinary);
            AssertNativeArtifactPortability(repoRoot, macBinary, windowsBinary);

            var nativeBinary = GetNativeBinary(macBinary, windowsBinary);
            if (nativeBinary == null)
            {
                var nativeRid = GetNativeRuntimeIdentifier();
                Assert.NotNull(nativeRid);
                nativeBinary = Publish(
                    repoRoot, nativeRid, Path.Combine(publishRoot, "native"));
            }

            AssertWordSmoke(nativeBinary, publishRoot);
            AssertAgentGuidanceSmoke(nativeBinary, publishRoot);
            AssertDefaultCliSurfaceStillWorks(nativeBinary, publishRoot);
        }
        finally
        {
            try { Directory.Delete(publishRoot, recursive: true); } catch { }

            // `dotnet publish -r` rewrites the referenced project's assets file for
            // the last RID. Restore the test graph so a later `--no-restore` build
            // still has the host runtime target available.
            var testProject = Path.Combine(repoRoot, "tests", "OfficeCli.Tests", "OfficeCli.Tests.csproj");
            var restore = RunProcess("dotnet", ["restore", testProject, "--nologo"], repoRoot,
                TimeSpan.FromMinutes(2));
            Assert.True(restore.ExitCode == 0,
                $"test project restore failed after publish smoke.\n{restore.Stderr}");
        }
    }

    private static string Publish(string repoRoot, string rid, string outputDirectory)
    {
        var project = Path.Combine(repoRoot, "src", "officecli", "officecli.csproj");
        var result = RunProcess(
            "dotnet",
            ["publish", project, "-c", "Release", "-r", rid, "--self-contained", "true",
             "-o", outputDirectory, "--nologo"],
            repoRoot,
            TimeSpan.FromMinutes(5));

        Assert.True(result.ExitCode == 0,
            $"publish {rid} failed.\nstdout:\n{result.Stdout}\nstderr:\n{result.Stderr}");

        var executable = rid.StartsWith("win-", StringComparison.Ordinal)
            ? Path.Combine(outputDirectory, "officecli.exe")
            : Path.Combine(outputDirectory, "officecli");
        Assert.True(File.Exists(executable), $"Missing published executable for {rid}: {executable}");
        return executable;
    }

    private static void AssertNativeArtifactPortability(
        string repoRoot,
        string macBinary,
        string windowsBinary)
    {
        ProcessResult? result = null;
        if (OperatingSystem.IsMacOS())
        {
            var verifier = Path.Combine(repoRoot, "build", "verify-macos-portability.sh");
            result = RunProcess("/bin/bash", [verifier, macBinary], repoRoot, TimeSpan.FromSeconds(30));
        }
        else if (OperatingSystem.IsWindows())
        {
            var verifier = Path.Combine(repoRoot, "build", "verify-windows-portability.ps1");
            result = RunProcess(
                "powershell.exe",
                ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", verifier, windowsBinary],
                repoRoot,
                TimeSpan.FromSeconds(30));
        }

        if (result != null)
            Assert.True(result.ExitCode == 0, $"published artifact portability failed.\n{result.Stderr}");
    }

    private static void AssertWordSmoke(string binary, string workingRoot)
    {
        var document = Path.Combine(workingRoot, "published-smoke.docx");
        var environment = new Dictionary<string, string>
        {
            ["OFFICECLI_NO_AUTO_RESIDENT"] = "1",
        };

        var create = RunProcess(binary, ["create", document, "--json"], workingRoot,
            TimeSpan.FromSeconds(30), environment);
        Assert.True(create.ExitCode == 0, create.Stderr);
        var add = RunProcess(binary,
            ["add", document, "/body", "--type", "paragraph", "--prop", "text=published smoke", "--json"],
            workingRoot, TimeSpan.FromSeconds(30), environment);
        Assert.True(add.ExitCode == 0, add.Stderr);

        var view = RunProcess(binary, ["view", document, "text", "--json"], workingRoot,
            TimeSpan.FromSeconds(30), environment);
        Assert.True(view.ExitCode == 0, view.Stderr);
        using (var viewDocument = JsonDocument.Parse(view.Stdout))
        {
            Assert.True(viewDocument.RootElement.GetProperty("success").GetBoolean(), view.Stdout);
            Assert.Contains("published smoke", viewDocument.RootElement.GetProperty("data").GetRawText());
        }

        var validate = RunProcess(binary, ["validate", document, "--json"], workingRoot,
            TimeSpan.FromSeconds(30), environment);
        Assert.True(validate.ExitCode == 0, validate.Stderr);
        using var validateDocument = JsonDocument.Parse(validate.Stdout);
        Assert.True(validateDocument.RootElement.GetProperty("success").GetBoolean(), validate.Stdout);
    }

    /// <summary>
    /// Phase-0 Agent guidance resources must survive single-file publish.
    /// </summary>
    private static void AssertAgentGuidanceSmoke(string binary, string workingRoot)
    {
        var environment = new Dictionary<string, string>
        {
            ["OFFICECLI_NO_AUTO_RESIDENT"] = "1",
        };

        var help = RunProcess(binary, ["help", "--surface", "agent-json"], workingRoot,
            TimeSpan.FromSeconds(30), environment);
        Assert.True(help.ExitCode == 0, help.Stderr);
        Assert.Contains("{{OFFICE_BATCH_TOOL}}", help.Stdout, StringComparison.Ordinal);
        Assert.Contains("{{OFFICE_RUN_TOOL}}", help.Stdout, StringComparison.Ordinal);
        Assert.DoesNotContain("--commands", help.Stdout, StringComparison.Ordinal);
        Assert.DoesNotContain("tool.file.office.", help.Stdout, StringComparison.Ordinal);

        var catalog = RunProcess(binary, ["load_skill", "--surface", "agent-json"], workingRoot,
            TimeSpan.FromSeconds(30), environment);
        Assert.True(catalog.ExitCode == 0, catalog.Stderr);
        foreach (var name in new[] { "word", "excel", "pptx", "word-form", "morph-ppt" })
            Assert.Contains($"## {name}", catalog.Stdout, StringComparison.Ordinal);

        foreach (var name in new[] { "word", "excel", "pptx", "word-form" })
        {
            var skill = RunProcess(binary, ["load_skill", name, "--surface", "agent-json"], workingRoot,
                TimeSpan.FromSeconds(30), environment);
            Assert.True(skill.ExitCode == 0, $"{name}: {skill.Stderr}");
            Assert.Contains("{{OFFICE_BATCH_TOOL}}", skill.Stdout, StringComparison.Ordinal);
            Assert.DoesNotContain("tool.file.office.", skill.Stdout, StringComparison.Ordinal);
        }

        var morph = RunProcess(binary, ["load_skill", "morph-ppt", "--surface", "agent-json"], workingRoot,
            TimeSpan.FromSeconds(30), environment);
        Assert.True(morph.ExitCode == 0, morph.Stderr);

        var morphRef = RunProcess(
            binary,
            ["load_skill", "morph-ppt", "--surface", "agent-json", "--path", "references/decision-rules.md"],
            workingRoot, TimeSpan.FromSeconds(30), environment);
        Assert.True(morphRef.ExitCode == 0, morphRef.Stderr);
        Assert.Contains("audience", morphRef.Stdout, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("officecli help", morphRef.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    private static void AssertDefaultCliSurfaceStillWorks(string binary, string workingRoot)
    {
        var environment = new Dictionary<string, string>
        {
            ["OFFICECLI_NO_AUTO_RESIDENT"] = "1",
        };

        var help = RunProcess(binary, ["help"], workingRoot, TimeSpan.FromSeconds(30), environment);
        Assert.True(help.ExitCode == 0, help.Stderr);
        Assert.Contains("Schema 参考", help.Stdout, StringComparison.Ordinal);
        Assert.DoesNotContain("{{OFFICE_BATCH_TOOL}}", help.Stdout, StringComparison.Ordinal);

        var skill = RunProcess(binary, ["load_skill", "word"], workingRoot,
            TimeSpan.FromSeconds(30), environment);
        Assert.True(skill.ExitCode == 0, skill.Stderr);
        Assert.DoesNotContain("{{OFFICE_BATCH_TOOL}}", skill.Stdout, StringComparison.Ordinal);
    }

    private static string? GetNativeBinary(string macBinary, string windowsBinary)
    {
        if (OperatingSystem.IsMacOS())
            return RuntimeInformation.ProcessArchitecture == Architecture.Arm64 ? macBinary : null;
        if (OperatingSystem.IsWindows())
            return RuntimeInformation.ProcessArchitecture == Architecture.X64 ? windowsBinary : null;
        return null;
    }

    private static string? GetNativeRuntimeIdentifier()
    {
        var architecture = RuntimeInformation.ProcessArchitecture switch
        {
            Architecture.X64 => "x64",
            Architecture.Arm64 => "arm64",
            _ => null,
        };
        if (architecture == null) return null;
        if (OperatingSystem.IsMacOS()) return $"osx-{architecture}";
        if (OperatingSystem.IsWindows()) return $"win-{architecture}";
        if (OperatingSystem.IsLinux()) return $"linux-{architecture}";
        return null;
    }

    private static ProcessResult RunProcess(
        string fileName,
        IEnumerable<string> arguments,
        string workingDirectory,
        TimeSpan timeout,
        IReadOnlyDictionary<string, string>? environment = null)
    {
        var startInfo = new ProcessStartInfo(fileName)
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        foreach (var argument in arguments)
            startInfo.ArgumentList.Add(argument);
        if (environment != null)
            foreach (var pair in environment)
                startInfo.Environment[pair.Key] = pair.Value;

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"Could not start {fileName}");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(timeout))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException($"{fileName} did not finish within {timeout}.");
        }

        return new ProcessResult(
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

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
