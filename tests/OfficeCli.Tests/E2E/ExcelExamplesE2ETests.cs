// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using FluentAssertions;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public sealed class ExcelExamplesE2ETests : ExcelTestBase
{
    // ==================== Top-level examples ====================

    [Fact]
    public void CellFormatting_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/cell-formatting.sh");
    }

    [Fact]
    public void Charts_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts.sh");
    }

    [Fact]
    public void ConditionalFormatting_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/conditional-formatting.sh");
    }

    [Fact]
    public void DataValidation_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/data-validation.sh");
    }

    [Fact]
    public void PivotTables_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/pivot-tables.sh");
    }

    [Fact]
    public void Shapes_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/shapes.sh");
    }

    [Fact]
    public void SheetSettings_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/sheet-settings.sh");
    }

    [Fact]
    public void Slicers_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/slicers.sh");
    }

    [Fact]
    public void Sparklines_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/sparklines.sh");
    }

    [Fact]
    public void WorkbookSettings_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/workbook-settings.sh");
    }

    // ==================== Charts sub-scripts ====================

    [Fact]
    public void ChartsAdvanced_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-advanced.sh");
    }

    [Fact]
    public void ChartsArea_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-area.sh");
    }

    [Fact]
    public void ChartsBar_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-bar.sh");
    }

    [Fact]
    public void ChartsBasic_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-basic.sh");
    }

    [Fact]
    public void ChartsBoxwhisker_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-boxwhisker.sh");
    }

    [Fact]
    public void ChartsBubble_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-bubble.sh");
    }

    [Fact]
    public void ChartsColumn_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-column.sh");
    }

    [Fact]
    public void ChartsCombo_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-combo.sh");
    }

    [Fact]
    public void ChartsExtended_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-extended.sh");
    }

    [Fact]
    public void ChartsHistogram_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-histogram.sh");
    }

    [Fact]
    public void ChartsLine_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-line.sh");
    }

    [Fact]
    public void ChartsPie_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-pie.sh");
    }

    [Fact]
    public void ChartsRadar_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-radar.sh");
    }

    [Fact]
    public void ChartsScatter_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-scatter.sh");
    }

    [Fact]
    public void ChartsStock_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-stock.sh");
    }

    [Fact]
    public void ChartsWaterfall_Script_RunsAndProducesValidWorkbook()
    {
        RunExampleScript("examples/excel/charts/charts-waterfall.sh");
    }

    // ==================== Helpers ====================

    private void RunExampleScript(string relativeScript)
    {
        var scriptPath = CopyExampleScriptToTemp(relativeScript);
        var workDir = Path.GetDirectoryName(scriptPath)!;

        // Put the built officecli on PATH so the script can find it.
        var exe = FindOfficeCliExecutable();
        var exeDir = exe != null
            ? Path.GetDirectoryName(exe)
            : Path.GetDirectoryName(FindOfficeCliAssembly());

        var env = exeDir != null
            ? new Dictionary<string, string>
            {
                ["PATH"] = $"{exeDir}:{Environment.GetEnvironmentVariable("PATH")}"
            }
            : null;

        var result = RunShellScriptOk(scriptPath, workDir, env);

        // Validate every generated workbook in the temp directory.
        var xlsxFiles = Directory.GetFiles(workDir, "*.xlsx");
        xlsxFiles.Should().NotBeEmpty(
            $"{relativeScript} should produce at least one .xlsx file in the work directory");

        foreach (var file in xlsxFiles)
        {
            TrackTempFile(file);
            RunCliOk("validate", file);
        }
    }
}
