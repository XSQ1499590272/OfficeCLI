// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelDumpBatchRenderTests : ExcelTestBase
{
    [Fact]
    public void CliCreateBatchDumpRawValidateAndRender_RoundTripExcelWorkbook()
    {
        var path = NewTempWorkbookPath();
        var replayPath = NewTempWorkbookPath();
        var dumpPath = NewTempPath(".json");
        var htmlPath = NewTempPath(".html");
        var pngPath = NewTempPath(".png");

        RunCliOk("create", path, "--force");

        const string batchJson = """
            [
              {"command":"add","parent":"/Sheet1/A1","type":"cell","props":{"value":"Cli Batch","bold":"true","fill":"D9EAD3"}},
              {"command":"add","parent":"/Sheet1/B1","type":"cell","props":{"value":"Amount"}},
              {"command":"add","parent":"/Sheet1/A2","type":"cell","props":{"value":"Jan"}},
              {"command":"add","parent":"/Sheet1/B2","type":"cell","props":{"value":"10","type":"number"}},
              {"command":"add","parent":"/Sheet1/A3","type":"cell","props":{"value":"Feb"}},
              {"command":"add","parent":"/Sheet1/B3","type":"cell","props":{"value":"15","type":"number"}},
              {"command":"add","parent":"/Sheet1/B4","type":"cell","props":{"formula":"SUM(B2:B3)"}},
              {"command":"view","mode":"html"},
              {"command":"raw","part":"/Sheet1"},
              {"command":"validate"}
            ]
            """;
        var batch = RunCliOk("batch", path, "--commands", batchJson, "--json");
        batch.Stdout.Should().Contain("\"success\": true");
        var batchDetails = ReadSpilledOutput(batch.Stdout);
        batchDetails.Should().Contain("Cli Batch");
        batchDetails.Should().Contain("Validation passed: no errors found.");

        using (var readOnly = OpenReadOnly(path))
        {
            ReadNode(readOnly, "/Sheet1/A1").Text.Should().Be("Cli Batch");
            ReadNode(readOnly, "/Sheet1/B4").Format.Should().Contain("formula", "SUM(B2:B3)");
            readOnly.Validate().Should().BeEmpty();
        }

        RunCliOk("dump", path, "/", "--out", dumpPath);
        var dumped = File.ReadAllText(dumpPath);
        dumped.Should().Contain("\"command\":\"import\"");
        dumped.Should().Contain("Cli Batch");

        RunCliOk("create", replayPath, "--force");
        RunCliOk("batch", replayPath, "--input", dumpPath, "--json");
        using (var replayed = OpenReadOnly(replayPath))
        {
            ReadNode(replayed, "/Sheet1/A1").Text.Should().Be("Cli Batch");
            ReadNode(replayed, "/Sheet1/B4").Format.Should().Contain("formula", "SUM(B2:B3)");
            replayed.Validate().Should().BeEmpty();
        }

        var rawSheet = RunCliOk("raw", path, "/Sheet1", "--start", "1", "--end", "2", "--cols", "A,B");
        rawSheet.Stdout.Should().Contain("Cli Batch");
        rawSheet.Stdout.Should().Contain("<x:sheetData");

        var rawWorkbook = RunCliOk("raw", path, "/workbook");
        rawWorkbook.Stdout.Should().Contain("<x:workbook");

        var validation = RunCliOk("validate", path, "--json");
        validation.Stdout.Should().Contain("\"success\": true");

        var html = RunCliOk("view", path, "html", "--out", htmlPath);
        html.Stdout.Should().Contain(Path.GetFullPath(htmlPath));
        File.Exists(htmlPath).Should().BeTrue();
        var htmlText = File.ReadAllText(htmlPath);
        htmlText.Should().Contain("<html");
        htmlText.Should().Contain("Cli Batch");

        var screenshot = RunCli("view", path, "screenshot", "--out", pngPath, "--screenshot-width", "480", "--screenshot-height", "320");
        if (screenshot.ExitCode == 0)
        {
            File.Exists(pngPath).Should().BeTrue();
            var png = File.ReadAllBytes(pngPath);
            png.Should().StartWith(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
            png.Length.Should().BeGreaterThan(100);
        }
        else
        {
            (screenshot.Stdout + screenshot.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

}
