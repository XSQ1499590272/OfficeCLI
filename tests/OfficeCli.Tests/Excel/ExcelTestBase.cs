// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Excel;

public abstract class ExcelTestBase : IDisposable
{
    private readonly List<string> _paths = new();

    protected string CreateWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_excel_{Guid.NewGuid():N}.xlsx");
        _paths.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    protected ExcelHandler OpenEditable(string path) => new(path, editable: true);

    protected ExcelHandler OpenReadOnly(string path) => new(path, editable: false);

    protected string TrackTempFile(string path)
    {
        _paths.Add(path);
        return path;
    }

    protected string TrackTempDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_excel_{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        _paths.Add(path);
        return path;
    }

    protected string NewTempWorkbookPath() => NewTempPath(".xlsx");

    protected string NewTempPath(string extension)
    {
        var path = TrackTempFile(Path.Combine(Path.GetTempPath(), $"officecli_excel_{Guid.NewGuid():N}{extension}"));
        if (File.Exists(path)) File.Delete(path);
        return path;
    }

    protected static DocumentNode ReadNode(ExcelHandler handler, string path)
    {
        var node = handler.Get(path);
        node.Should().NotBeNull(path);
        return node!;
    }

    protected static IReadOnlyList<DocumentNode> Query(ExcelHandler handler, string selector)
        => handler.Query(selector);

    protected static string ReadZipEntry(string workbookPath, string entryName)
    {
        using var archive = ZipFile.OpenRead(workbookPath);
        var entry = archive.GetEntry(entryName);
        entry.Should().NotBeNull($"{workbookPath} should contain {entryName}");
        using var stream = entry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    protected CliResult RunCliOk(params string[] args)
    {
        var result = RunCli(args);
        result.ExitCode.Should().Be(0, $"stdout:\n{result.Stdout}\nstderr:\n{result.Stderr}");
        return result;
    }

    protected static CliResult RunCli(params string[] args)
    {
        var exe = FindOfficeCliExecutable();
        var psi = new ProcessStartInfo(exe ?? "dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        psi.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";
        if (exe == null)
            psi.ArgumentList.Add(FindOfficeCliAssembly());
        foreach (var arg in args)
            psi.ArgumentList.Add(arg);

        using var process = Process.Start(psi);
        process.Should().NotBeNull();
        var stdout = process!.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(60_000))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException($"officecli timed out: {string.Join(" ", args)}");
        }

        return new CliResult(process.ExitCode, stdout, stderr);
    }

    protected string CopyExampleScriptToTemp(string relativeScriptPath)
    {
        var source = Path.Combine(RepoRoot(), relativeScriptPath);
        var dir = TrackTempDirectory();
        var script = Path.Combine(dir, Path.GetFileName(source));
        File.Copy(source, script);
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(script, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        return script;
    }

    protected CliResult RunShellScriptOk(string scriptPath, string workingDirectory, IReadOnlyDictionary<string, string>? env = null)
    {
        var psi = new ProcessStartInfo("bash", scriptPath)
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        psi.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";
        if (env != null)
        {
            foreach (var pair in env)
                psi.Environment[pair.Key] = pair.Value;
        }

        using var process = Process.Start(psi);
        process.Should().NotBeNull();
        var stdout = process!.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(180_000))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException($"{scriptPath} timed out");
        }

        var result = new CliResult(process.ExitCode, stdout, stderr);
        result.ExitCode.Should().Be(0, $"stdout:\n{stdout}\nstderr:\n{stderr}");
        return result;
    }

    protected string CreateTinyPng(string path)
    {
        File.WriteAllBytes(path, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
        return TrackTempFile(path);
    }

    protected string CreateOlePayload(string path)
    {
        File.WriteAllText(path, "embedded payload");
        return TrackTempFile(path);
    }

    protected string WriteCsv(string path, params string[] rows)
    {
        File.WriteAllText(path, string.Join(Environment.NewLine, rows) + Environment.NewLine);
        return TrackTempFile(path);
    }

    protected static string ReadSpilledOutput(string stdout)
    {
        var match = Regex.Match(stdout, @"""outputFile""\s*:\s*""(?<path>[^""]+)""");
        if (!match.Success) return stdout;
        var path = match.Groups["path"].Value;
        return File.Exists(path) ? File.ReadAllText(path) : stdout;
    }

    protected static string RepoRoot()
    {
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir != null; dir = dir.Parent)
            if (File.Exists(Path.Combine(dir.FullName, "officecli.slnx"))
                && Directory.Exists(Path.Combine(dir.FullName, "examples")))
                return dir.FullName;
        return Directory.GetCurrentDirectory();
    }

    protected static string? FindOfficeCliExecutable()
    {
        var assembly = FindOfficeCliAssembly();
        var dir = Path.GetDirectoryName(assembly);
        if (dir == null) return null;
        var name = OperatingSystem.IsWindows() ? "officecli.exe" : "officecli";
        var exe = Path.Combine(dir, name);
        return File.Exists(exe) ? exe : null;
    }

    protected static string FindOfficeCliAssembly()
    {
        var copiedAssembly = typeof(ExcelHandler).Assembly.Location;
        if (HasHostPolicy(Path.GetDirectoryName(copiedAssembly)))
            return copiedAssembly;

        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir != null; dir = dir.Parent)
        {
            var binDir = Path.Combine(dir.FullName, "src", "officecli", "bin");
            if (!Directory.Exists(binDir)) continue;
            var ridSegment = $"{Path.DirectorySeparatorChar}{RuntimeInformation.RuntimeIdentifier}{Path.DirectorySeparatorChar}";

            var candidate = Directory.EnumerateFiles(binDir, "officecli.dll", SearchOption.AllDirectories)
                .Where(path => path.Contains(ridSegment) && HasHostPolicy(Path.GetDirectoryName(path)))
                .OrderBy(path => path.Contains($"{Path.DirectorySeparatorChar}Debug{Path.DirectorySeparatorChar}") ? 0 : 1)
                .ThenBy(path => path.Length)
                .FirstOrDefault();
            if (candidate != null)
                return candidate;
        }

        return copiedAssembly;
    }

    private static bool HasHostPolicy(string? directory)
        => directory != null
           && Directory.Exists(directory)
           && Directory.EnumerateFiles(directory, "libhostpolicy.*").Any();

    public void Dispose()
    {
        foreach (var path in _paths)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
                else if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
            }
            catch
            {
                // Best effort cleanup for temp files held by a failed assertion.
            }
        }
    }

    protected sealed record CliResult(int ExitCode, string Stdout, string Stderr);
}
