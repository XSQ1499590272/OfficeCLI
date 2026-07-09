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

namespace OfficeCli.Tests.Pptx;

public abstract class PptTestBase : IDisposable
{
    private readonly List<string> _paths = new();

    protected string CreatePresentation()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_ppt_{Guid.NewGuid():N}.pptx");
        _paths.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    protected PowerPointHandler OpenEditable(string path) => new(path, editable: true);

    protected PowerPointHandler OpenReadOnly(string path) => new(path, editable: false);

    protected string TrackTempFile(string path)
    {
        _paths.Add(path);
        return path;
    }

    protected string TrackTempDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_ppt_{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        _paths.Add(path);
        return path;
    }

    protected string NewTempPresentationPath() => NewTempPath(".pptx");

    protected string NewTempPath(string extension)
    {
        var path = TrackTempFile(Path.Combine(Path.GetTempPath(), $"officecli_ppt_{Guid.NewGuid():N}{extension}"));
        if (File.Exists(path)) File.Delete(path);
        return path;
    }

    protected static string ReadZipEntry(string pptxPath, string entryName)
    {
        using var archive = ZipFile.OpenRead(pptxPath);
        var entry = archive.GetEntry(entryName);
        entry.Should().NotBeNull($"{pptxPath} should contain {entryName}");
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

    protected string CreateTinySvg(string path)
    {
        File.WriteAllText(path, @"<?xml version=""1.0"" encoding=""UTF-8""?>
<svg xmlns=""http://www.w3.org/2000/svg"" width=""10"" height=""10"">
  <rect width=""10"" height=""10"" fill=""red""/>
</svg>");
        return TrackTempFile(path);
    }

    protected string CreateOlePayload(string path)
    {
        File.WriteAllText(path, "embedded payload");
        return TrackTempFile(path);
    }

    /// <summary>
    /// Creates a minimal valid glTF 2.0 binary (.glb) file at the given path.
    /// Contains a single-node scene with no meshes — the smallest valid GLB.
    /// </summary>
    protected string CreateTinyGlb(string path)
    {
        // Minimal glTF 2.0 JSON: a scene with a single node, no buffers.
        var json = @"{""asset"":{""version"":""2.0""},""scene"":0,""scenes"":[{""nodes"":[0]}],""nodes"":[{""name"":""empty""}],""meshes"":[]}";
        var jsonBytes = Encoding.UTF8.GetBytes(json);

        // Pad JSON chunk to 4-byte alignment with space chars (0x20).
        while (jsonBytes.Length % 4 != 0)
        {
            Array.Resize(ref jsonBytes, jsonBytes.Length + 1);
            jsonBytes[^1] = 0x20;
        }

        // GLB header (12 bytes) + JSON chunk header (8 bytes) + JSON chunk data
        var totalLength = 12 + 8 + jsonBytes.Length;
        var glb = new byte[totalLength];

        // Magic: "glTF"
        glb[0] = 0x67; glb[1] = 0x6C; glb[2] = 0x54; glb[3] = 0x46;
        // Version: 2
        BitConverter.GetBytes((uint)2).CopyTo(glb, 4);
        // Total length (little-endian)
        BitConverter.GetBytes((uint)totalLength).CopyTo(glb, 8);
        // JSON chunk length
        BitConverter.GetBytes((uint)jsonBytes.Length).CopyTo(glb, 12);
        // JSON chunk type: "JSON"
        glb[16] = 0x4A; glb[17] = 0x53; glb[18] = 0x4F; glb[19] = 0x4E;
        // JSON chunk data
        Array.Copy(jsonBytes, 0, glb, 20, jsonBytes.Length);

        File.WriteAllBytes(path, glb);
        return TrackTempFile(path);
    }

    /// <summary>
    /// Creates a minimal valid MP4 file at the given path.
    /// Contains only an ftyp box — the smallest valid MP4 recognizable by media players.
    /// </summary>
    protected string CreateTinyVideo(string path)
    {
        // Minimal MP4 with ftyp box only (isom, minor version 512, compatible brands: isom/iso2/mp41).
        var mp4 = Convert.FromBase64String(
            "AAAAGGZ0eXBpc29tAAAgAAABaWlzb202AAAGAAAAbXA0MQ==");
        File.WriteAllBytes(path, mp4);
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
        var copiedAssembly = typeof(PowerPointHandler).Assembly.Location;
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
