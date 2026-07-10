// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using FluentAssertions;

using OfficeCli.Tests.Pptx;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public sealed class PptExamplesE2ETests : PptTestBase
{
    // ==================== Top-level examples ====================

    [Fact]
    public void Animations_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/animations.sh");
    }

    [Fact]
    public void Diagram_Script_RunsAndProducesValidPresentation()
    {
        // diagram.sh references pie.mmd in the same source directory.
        RunExampleScript("examples/ppt/diagram.sh",
            extraFiles: new[] { "examples/ppt/pie.mmd" });
    }

    [Fact]
    public void Presentation_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/presentation.sh");
    }

    [Fact]
    public void PresentationSettings_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/presentation-settings.sh");
    }

    // ==================== Charts ====================

    [Fact]
    public void Charts3d_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-3d.sh");
    }

    [Fact]
    public void ChartsAdvanced_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-advanced.sh");
    }

    [Fact]
    public void ChartsArea_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-area.sh");
    }

    [Fact]
    public void ChartsBar_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-bar.sh");
    }

    [Fact]
    public void ChartsBubble_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-bubble.sh");
    }

    [Fact]
    public void ChartsColumn_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-column.sh");
    }

    [Fact]
    public void ChartsCombo_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-combo.sh");
    }

    [Fact]
    public void ChartsDoughnut_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-doughnut.sh");
    }

    [Fact]
    public void ChartsLine_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-line.sh");
    }

    [Fact]
    public void ChartsPie_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-pie.sh");
    }

    [Fact]
    public void ChartsRadar_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-radar.sh");
    }

    [Fact]
    public void ChartsScatter_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-scatter.sh");
    }

    [Fact]
    public void ChartsStock_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-stock.sh");
    }

    [Fact]
    public void ChartsWaterfall_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/charts/charts-waterfall.sh");
    }

    // ==================== Shapes ====================

    [Fact]
    public void ShapesBasic_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/shapes/shapes-basic.sh");
    }

    [Fact]
    public void ShapesConnectors_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/shapes/shapes-connectors.sh");
    }

    [Fact]
    public void ShapesEffects_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/shapes/shapes-effects.sh");
    }

    [Fact]
    public void ShapesTypography_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/shapes/shapes-typography.sh");
    }

    // ==================== Tables ====================

    [Fact]
    public void TablesBasic_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/tables/tables-basic.sh");
    }

    [Fact]
    public void TablesBorders_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/tables/tables-borders.sh");
    }

    [Fact]
    public void TablesFinancial_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/tables/tables-financial.sh");
    }

    [Fact]
    public void TablesMerged_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/tables/tables-merged.sh");
    }

    [Fact]
    public void TablesNested_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/tables/tables-nested.sh");
    }

    [Fact]
    public void TablesRowsCols_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/tables/tables-rows-cols.sh");
    }

    [Fact]
    public void TablesStyled_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/tables/tables-styled.sh");
    }

    // ==================== Textboxes ====================

    [Fact]
    public void TextboxesAdvanced_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/textboxes/textboxes-advanced.sh");
    }

    [Fact]
    public void TextboxesBasic_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/textboxes/textboxes-basic.sh");
    }

    // ==================== Transitions ====================

    [Fact]
    public void TransitionsBands_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-bands.sh");
    }

    [Fact]
    public void TransitionsBasic_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-basic.sh");
    }

    [Fact]
    public void TransitionsDirectional_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-directional.sh");
    }

    [Fact]
    public void TransitionsDynamic_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-dynamic.sh");
    }

    [Fact]
    public void TransitionsModern_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-modern.sh");
    }

    [Fact]
    public void TransitionsMorph_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-morph.sh");
    }

    [Fact]
    public void TransitionsRandom_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-random.sh");
    }

    [Fact]
    public void TransitionsShapes_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-shapes.sh");
    }

    [Fact]
    public void TransitionsTiming_Script_RunsAndProducesValidPresentation()
    {
        RunExampleScript("examples/ppt/transitions/transitions-timing.sh");
    }

    // ==================== Quarantined: external dependencies ====================

    // Quarantined: pictures-basic.sh — requires Pillow (pip install Pillow) to
    // synthesize sample PNG images at runtime via embedded Python. The script
    // cannot produce its output without those generated image files.

    // Quarantined: ole-embed.sh — requires Pillow (pip install Pillow) to
    // synthesize preview thumbnail PNGs for embedded OLE objects. While the
    // script builds .xlsx/.docx payloads with officecli itself, the thumbnail
    // generation step depends on Python/Pillow.

    // Quarantined: video.sh — requires imageio, imageio-ffmpeg, and numpy
    // (pip install imageio imageio-ffmpeg numpy) to generate a demo MP4 video
    // and cover image via embedded Python. Without those packages the script
    // exits with an error before producing any output.

    // Quarantined: 3d-model.sh — requires a 4.3 MB sun.glb model file in
    // examples/ppt/models/. While the file exists in the repo, copying large
    // binary assets to a temp directory for every test run is impractical.
    // This script is excluded from automated E2E testing.

    // Quarantined: templates/**/*.sh — each template build script operates on a
    // pre-existing template .pptx file and applies style overrides. These are
    // style-demonstration scripts, not self-contained examples. They depend on
    // hand-crafted template files with specific slide layouts and content.

    // ==================== Helpers ====================

    /// <summary>
    /// Copies an example shell script (and optional extra files) to a temp
    /// directory, puts the built officecli on PATH, runs the script, and
    /// validates every generated .pptx file.
    /// </summary>
    /// <param name="relativeScript">Path relative to repo root, e.g. "examples/ppt/presentation.sh".</param>
    /// <param name="extraFiles">Additional files (relative to repo root) to copy alongside the script.</param>
    private void RunExampleScript(string relativeScript, string[]? extraFiles = null)
    {
        var scriptPath = CopyExampleScriptToTemp(relativeScript);
        var workDir = Path.GetDirectoryName(scriptPath)!;

        // Copy any extra files the script needs at runtime (e.g. .mmd, model files).
        if (extraFiles != null)
        {
            foreach (var file in extraFiles)
            {
                var source = Path.Combine(RepoRoot(), file);
                var dest = Path.Combine(workDir, Path.GetFileName(file));
                File.Copy(source, dest);
            }
        }

        // Put the built officecli on PATH so scripts can find it.
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

        // Validate every generated presentation in the temp directory.
        var pptxFiles = Directory.GetFiles(workDir, "*.pptx");
        pptxFiles.Should().NotBeEmpty(
            $"{relativeScript} should produce at least one .pptx file in the work directory");

        foreach (var file in pptxFiles)
        {
            TrackTempFile(file);
            RunCliOk("validate", file);

            // Smoke readback: verify slide count via 'get'.
            var getResult = RunCli("get", file, "/");
            getResult.ExitCode.Should().Be(0,
                $"officecli get on {Path.GetFileName(file)} should succeed.\nstderr:\n{getResult.Stderr}");

            // Verify stdout contains real content — not just an empty success.
            getResult.Stdout.Should().NotBeNullOrWhiteSpace(
                $"officecli get on {Path.GetFileName(file)} should return content.");
            getResult.Stdout!.Should().Contain("/slide[",
                $"officecli get on {Path.GetFileName(file)} should reference at least one slide node.");

            // Log the slide count for diagnostics.
            var slideCount = 0;
            var idx = 0;
            while ((idx = getResult.Stdout.IndexOf("/slide[", idx, StringComparison.Ordinal)) != -1)
            {
                slideCount++;
                idx++;
            }

            Debug.WriteLine($"[PptExamplesE2E] {Path.GetFileName(file)}: {slideCount} slide(s) detected.");
        }
    }
}
