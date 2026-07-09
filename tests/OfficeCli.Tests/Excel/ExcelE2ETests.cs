// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using FluentAssertions;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelE2ETests : ExcelTestBase
{
    [Fact]
    public void CliExcel_CreateOpenBatchSaveCloseValidate()
    {
        var path = NewTempWorkbookPath();

        RunCliOk("create", path, "--force");
        RunCliOk("open", path);
        RunCliOk("batch", path, "--commands", """
            [
              {"command":"add","parent":"/Sheet1/A1","type":"cell","props":{"value":"Region"}},
              {"command":"add","parent":"/Sheet1/B1","type":"cell","props":{"value":"Sales"}},
              {"command":"add","parent":"/Sheet1/A2","type":"cell","props":{"value":"West"}},
              {"command":"add","parent":"/Sheet1/B2","type":"cell","props":{"value":"1200","type":"number"}},
              {"command":"add","parent":"/Sheet1/B3","type":"cell","props":{"formula":"SUM(B2:B2)"}}
            ]
            """, "--json");
        RunCliOk("save", path);
        RunCliOk("close", path);
        RunCliOk("validate", path);

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/B3").Format.Should().Contain("formula", "SUM(B2:B2)");
    }

    [Fact]
    public void CliExcel_DumpBatchRawReplay()
    {
        var source = NewTempWorkbookPath();
        var target = NewTempWorkbookPath();
        var dump = NewTempPath(".json");

        RunCliOk("create", source, "--force");
        RunCliOk("batch", source, "--commands", """
            [
              {"command":"add","parent":"/Sheet1/A1","type":"cell","props":{"value":"Item"}},
              {"command":"add","parent":"/Sheet1/B1","type":"cell","props":{"value":"Qty"}},
              {"command":"add","parent":"/Sheet1/A2","type":"cell","props":{"value":"Pen"}},
              {"command":"add","parent":"/Sheet1/B2","type":"cell","props":{"value":"2","type":"number"}},
              {"command":"add","parent":"/Sheet1/B3","type":"cell","props":{"formula":"SUM(B2:B2)"}}
            ]
            """);
        RunCliOk("dump", source, "/", "--out", dump);
        RunCliOk("create", target, "--force");
        RunCliOk("batch", target, "--input", dump, "--json");

        var rawWorkbook = RunCliOk("raw", target, "/workbook");
        rawWorkbook.Stdout.Should().Contain("<x:workbook");
        var rawSheet = RunCliOk("raw", target, "/Sheet1");
        rawSheet.Stdout.Should().Contain("Pen").And.Contain("SUM(B2:B2)");
        RunCliOk("validate", target);
    }

    [Fact]
    public void CliExcel_ViewHtmlAlwaysRendersWorkbook()
    {
        var path = NewTempWorkbookPath();
        var htmlPath = NewTempPath(".html");

        RunCliOk("create", path, "--force");
        RunCliOk("batch", path, "--commands", """
            [
              {"command":"add","parent":"/Sheet1/A1","type":"cell","props":{"value":"Label","bold":"true"}},
              {"command":"add","parent":"/Sheet1/B1","type":"cell","props":{"value":"Very visible content"}},
              {"command":"set","path":"/Sheet1/col[A]","props":{"width":"18"}},
              {"command":"set","path":"/Sheet1/col[B]","props":{"width":"28"}}
            ]
            """);
        RunCliOk("view", path, "html", "--out", htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html").And.Contain("Very visible content").And.NotContain("###");
    }

    [Fact]
    public void CliExcel_ViewScreenshotRendersOrReportsMissingBackend()
    {
        var path = NewTempWorkbookPath();
        var pngPath = NewTempPath(".png");

        RunCliOk("create", path, "--force");
        RunCliOk("batch", path, "--commands", """
            [
              {"command":"add","parent":"/Sheet1/A1","type":"cell","props":{"value":"Screenshot"}},
              {"command":"add","parent":"/Sheet1/B1","type":"cell","props":{"value":"Smoke"}}
            ]
            """);

        var result = RunCli("view", path, "screenshot", "--out", pngPath, "--screenshot-width", "480", "--screenshot-height", "320");
        if (result.ExitCode == 0)
        {
            File.ReadAllBytes(pngPath).Should().StartWith(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
        }
        else
        {
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    [Fact]
    public void CliExcel_ExamplesSmoke()
    {
        foreach (var relativeScript in new[]
        {
            "examples/excel/cell-formatting.sh",
            "examples/excel/data-validation.sh",
            "examples/excel/pivot-tables.sh",
            "examples/excel/slicers.sh",
            "examples/excel/charts/charts-basic.sh",
            "examples/excel/charts/charts-extended.sh"
        })
        {
            var output = RunExampleScript(relativeScript);
            RunCliOk("validate", output);
        }
    }

    private string RunExampleScript(string relativeScript)
    {
        var source = Path.Combine(RepoRoot(), relativeScript);
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"officecli_example_{Guid.NewGuid():N}")).FullName;
        TrackTempFile(dir);
        var script = Path.Combine(dir, Path.GetFileName(source));
        File.Copy(source, script);

        var binDir = Directory.CreateDirectory(Path.Combine(dir, "bin")).FullName;
        var wrapper = Path.Combine(binDir, "officecli");
        var exe = FindOfficeCliExecutable();
        var command = exe != null
            ? $"exec \"{exe}\" \"$@\""
            : $"exec dotnet \"{FindOfficeCliAssembly()}\" \"$@\"";
        File.WriteAllText(wrapper, $"#!/bin/sh\n{command}\n");
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(wrapper, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);

        var psi = new ProcessStartInfo("bash", script)
        {
            WorkingDirectory = dir,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        psi.Environment["PATH"] = $"{binDir}{Path.PathSeparator}{psi.Environment["PATH"]}";
        psi.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";

        using var process = Process.Start(psi);
        process.Should().NotBeNull();
        var stdout = process!.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(180_000))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException($"{relativeScript} timed out");
        }
        process.ExitCode.Should().Be(0, $"stdout:\n{stdout}\nstderr:\n{stderr}");

        var output = Path.Combine(dir, Path.GetFileNameWithoutExtension(source) + ".xlsx");
        File.Exists(output).Should().BeTrue($"{relativeScript} should create {output}");
        TrackTempFile(output);
        return output;
    }
}
