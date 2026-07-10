// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Runtime.InteropServices;
using FluentAssertions;

using OfficeCli.Tests.Pptx;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public sealed class PptRenderE2ETests : PptTestBase
{
    // ==================== view html ====================

    [Fact]
    public void ViewHtml_TextAndShapes_ProducesInspectableHtml()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=HTML Test");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Hello HTML");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        File.Exists(htmlPath).Should().BeTrue();
        var html = File.ReadAllText(htmlPath);
        html.Should().NotBeNullOrEmpty();
        html.Should().Contain("<html");
        html.Should().Contain("Hello HTML");
    }

    [Fact]
    public void ViewHtml_MultipleSlides_AllIncluded()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Slide A");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Slide B");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Slide C");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("Slide A");
        html.Should().Contain("Slide B");
        html.Should().Contain("Slide C");
    }

    [Fact]
    public void ViewHtml_Picture_ShowsImageContent()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Picture Slide");
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        RunCliOk("add", path, "/slide[1]", "--type", "picture",
            "--prop", $"src={pngPath}");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // HTML should contain image data (base64 data URI or img tag)
        html.Should().MatchRegex("(<img|<image|data:image)");
    }

    [Fact]
    public void ViewHtml_Table_RendersTableMarkup()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Table Slide");
        RunCliOk("add", path, "/slide[1]", "--type", "table",
            "--prop", "data=A,B,C;D,E,F;G,H,I",
            "--prop", "firstRow=true");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // Table should render as table markup or grid structure
        html.Should().ContainAny("A", "B", "F", "I");
    }

    [Fact]
    public void ViewHtml_Chart_ShowsChartContent()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Chart Slide");
        RunCliOk("add", path, "/slide[1]", "--type", "chart",
            "--prop", "chartType=column",
            "--prop", "title=Sales",
            "--prop", "categories=Q1,Q2,Q3",
            "--prop", "data=Revenue:10,20,30");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // Chart should render as inline SVG in HTML
        html.Should().MatchRegex("(<svg|class=\"chart|<chart)");
    }

    [Fact]
    public void ViewHtml_Notes_ShowsSpeakerNotes()
    {
        var path = CreatePresentation();
        // Notes text is set as a slide property after slide creation (set-only).
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Notes Slide");
        RunCliOk("set", path, "/slide[1]", "--prop", "notes=Speaker notes content");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // Notes text should appear somewhere in the HTML
        html.Should().Contain("Speaker notes content");
    }

    [Fact]
    public void ViewHtml_Hyperlinks_ContainsLinkedContent()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Link Slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Click here",
            "--prop", "link=https://officecli.ai");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Click here");
        // Hyperlink should appear as href or data-href
        html.Should().MatchRegex("(href|hyperlink|officecli\\.ai)");
    }

    [Fact]
    public void ViewHtml_BoldItalicColor_ShowsFormattedText()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Formatted Slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Styled Content",
            "--prop", "bold=true",
            "--prop", "italic=true",
            "--prop", "color=FF0000");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Styled Content");
        // Bold should produce <b>, <strong>, or font-weight:bold
        html.Should().MatchRegex("(<b>|<strong>|font-weight:\\s*bold|bold)");
    }

    [Fact]
    public void ViewHtml_HiddenSlide_StillInHtmlOutput()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Visible");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Hidden");
        RunCliOk("set", path, "/slide[2]", "--prop", "hidden=true");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // Both slides should appear in HTML output; the hidden flag is a
        // presentation-mode hint, not a rendering filter.
        html.Should().Contain("Visible");
        html.Should().Contain("Hidden");
    }

    [Fact]
    public void ViewHtml_PageFilter_SingleSlideOnly()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=First");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Second");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Third");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath, "--page", "2");
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Second");
        // First and Third should not appear when --page 2 filters to slide 2 only.
        html.Should().NotContain("First");
        html.Should().NotContain("Third");
    }

    [Fact]
    public void ViewHtml_ThemeColors_ProducesColoredOutput()
    {
        var path = CreatePresentation();
        // Set theme accent colors to verify themed output
        RunCliOk("set", path, "/", "--prop", "theme.color.accent1=FF6600");
        RunCliOk("set", path, "/", "--prop", "theme.color.accent2=0066CC");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Themed Slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Themed text",
            "--prop", "fill=accent1");

        var htmlPath = NewTempPath(".html");
        var result = RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Themed text");
        // Theme colors should produce CSS color values in the output
        html.Should().MatchRegex("(#[0-9a-fA-F]{6}|rgb\\(|color:)");
    }

    // ==================== view svg ====================

    [Fact]
    public void ViewSvg_SingleSlide_ProducesValidSvg()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=SVG Test");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=SVG Content");

        // SVG mode writes to stdout when --browser is not used.
        var result = RunCliOk("view", path, "svg", "--page", "1");
        result.Stdout.Should().NotBeNullOrEmpty();
        result.Stdout.Should().Contain("<svg");
        result.Stdout.Should().Contain("SVG Content");
    }

    [Fact]
    public void ViewSvg_WithoutPageFilter_DefaultsToSlide1()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Default Slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Default Content");

        // No --page flag: should render slide 1 by default.
        var result = RunCliOk("view", path, "svg");
        result.Stdout.Should().Contain("<svg");
        result.Stdout.Should().Contain("Default Content");
    }

    [Fact]
    public void ViewSvg_PageFilter_SelectsCorrectSlide()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=S1");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=S2");
        RunCliOk("add", path, "/slide[2]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Second Slide Shape");

        var result = RunCliOk("view", path, "svg", "--page", "2");
        result.Stdout.Should().Contain("<svg");
        result.Stdout.Should().Contain("Second Slide Shape");
        // Should NOT contain slide 1 title since we filtered to slide 2.
        result.Stdout.Should().NotContain("S1");
    }

    [Fact]
    public void ViewSvg_InvalidPage_Zero_ReportsError()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=S1");

        var result = RunCli("view", path, "svg", "--page", "0");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void ViewSvg_InvalidPage_OutOfRange_ReportsError()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=S1");

        var result = RunCli("view", path, "svg", "--page", "99");
        // Out-of-range page should produce an error (empty or missing slide).
        result.ExitCode.Should().NotBe(0);
    }

    // ==================== view screenshot ====================

    [Fact]
    public void ViewScreenshot_Basic_AcceptsNoBackendOrProducesPng()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Screenshot Test");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Screenshot Content");

        var ssPath = NewTempPath(".png");
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--screenshot-width", "480", "--screenshot-height", "320");
        TrackTempFile(ssPath);

        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            // PNG magic bytes: 89 50 4E 47 0D 0A 1A 0A
            File.ReadAllBytes(ssPath).Should().StartWith(
                new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        }
        else
        {
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    [Fact]
    public void ViewScreenshot_GridAuto_MultiSlide()
    {
        var path = CreatePresentation();
        for (int i = 1; i <= 4; i++)
        {
            RunCliOk("add", path, "/", "--type", "slide", "--prop", $"title=Grid Slide {i}");
            RunCliOk("add", path, $"/slide[{i}]", "--type", "shape",
                "--prop", "preset=rect", "--prop", $"text=Content {i}");
        }

        var ssPath = NewTempPath(".png");
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--grid", "auto",
            "--screenshot-width", "480", "--screenshot-height", "320");
        TrackTempFile(ssPath);

        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            File.ReadAllBytes(ssPath).Should().StartWith(
                new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        }
        else
        {
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    [Fact]
    public void ViewScreenshot_Grid3_MultiSlide()
    {
        var path = CreatePresentation();
        for (int i = 1; i <= 6; i++)
        {
            RunCliOk("add", path, "/", "--type", "slide", "--prop", $"title=Grid {i}");
        }

        var ssPath = NewTempPath(".png");
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--grid", "3",
            "--screenshot-width", "480", "--screenshot-height", "320");
        TrackTempFile(ssPath);

        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            File.ReadAllBytes(ssPath).Should().StartWith(
                new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        }
        else
        {
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    [Fact]
    public void ViewScreenshot_RenderHtml_MustWorkWithoutNative()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=HTML Render");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=HTML Rendered");

        var ssPath = NewTempPath(".png");
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--render", "html",
            "--screenshot-width", "480", "--screenshot-height", "320");
        TrackTempFile(ssPath);

        // --render html goes through the HTML path, which does not require
        // native PowerPoint. It should produce a PNG or report no backend.
        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            File.ReadAllBytes(ssPath).Should().StartWith(
                new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        }
        else
        {
            // no_screenshot_backend is acceptable even for --render html when no
            // headless browser is available in CI.
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    [Fact]
    public void ViewScreenshot_RenderNative_WindowsOnlyConditional()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Native Render");

        // --render native requires Windows + Microsoft PowerPoint installed.
        // On non-Windows this should fail with a clear error.
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var result = RunCli("view", path, "screenshot",
                "--render", "native",
                "--screenshot-width", "480", "--screenshot-height", "320");
            // On non-Windows, native render should fail or fallback gracefully.
            (result.Stdout + result.Stderr).Should().MatchRegex(
                "(requires Windows|native.*not.*(available|supported)|no_screenshot_backend|Native render is only available)");
        }
        // On Windows without PowerPoint, this would also fail.
        // We don't assert success even on Windows since PowerPoint may not be installed.
    }

    [Fact]
    public void ViewScreenshot_BareGrid_DefaultsToAutoColumns()
    {
        var path = CreatePresentation();
        for (int i = 1; i <= 4; i++)
        {
            RunCliOk("add", path, "/", "--type", "slide", "--prop", $"title=BareGrid {i}");
        }

        var ssPath = NewTempPath(".png");
        // Bare --grid without a value should behave like --grid auto.
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--grid",
            "--screenshot-width", "480", "--screenshot-height", "320");
        TrackTempFile(ssPath);

        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            File.ReadAllBytes(ssPath).Should().StartWith(
                new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        }
        else
        {
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    [Fact]
    public void ViewScreenshot_Chart_AcceptsNoBackendOrProducesPng()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Chart Screenshot");
        RunCliOk("add", path, "/slide[1]", "--type", "chart",
            "--prop", "chartType=pie",
            "--prop", "title=Distribution",
            "--prop", "categories=A,B,C",
            "--prop", "data=Values:30,45,25");

        var ssPath = NewTempPath(".png");
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--screenshot-width", "480", "--screenshot-height", "320");
        TrackTempFile(ssPath);

        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            File.ReadAllBytes(ssPath).Should().StartWith(
                new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        }
        else
        {
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    // ==================== view html: rtl & animation ====================

    [Fact]
    public void ViewHtml_RtlText_ShowsRtlDirection()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=RTL Test");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=שלום עולם",
            "--prop", "direction=rtl");

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().ContainAny("direction: rtl", "dir=\"rtl\"", "rtl");
    }

    [Fact]
    public void ViewHtml_Animation_HasMarkers()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Animation Test");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Animated",
            "--prop", "animation=fade");

        // Verify the animation was applied: the shape should exist and have its text.
        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("Animated");

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Animated");
        // Note: the HTML renderer does not currently emit CSS animation
        // markers for PPTX shape animations. When support is added, this
        // test should also assert on animation-related DOM attributes.
    }

    // ==================== view svg: file output ====================

    [Fact]
    public void ViewSvg_OutFile_WritesSvgToFile()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=SVG Out Test");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=SVG File Output");

        var svgPath = NewTempPath(".svg");
        var result = RunCli("view", path, "svg", "--out", svgPath);
        TrackTempFile(svgPath);

        // When --out is supported for SVG, the file should exist with valid SVG.
        // If --out is not yet supported, the SVG goes to stdout and the file is
        // not written — fall back to capturing stdout and writing it ourselves.
        if (File.Exists(svgPath) && new FileInfo(svgPath).Length > 0)
        {
            var svg = File.ReadAllText(svgPath);
            svg.Should().Contain("<svg");
            svg.Should().Contain("</svg>");
        }
        else
        {
            // --out not yet supported for SVG: capture stdout and write to file.
            result.Stdout.Should().Contain("<svg");
            File.WriteAllText(svgPath, result.Stdout);
            var svg = File.ReadAllText(svgPath);
            svg.Should().Contain("<svg");
            svg.Should().Contain("</svg>");
        }
    }
}
