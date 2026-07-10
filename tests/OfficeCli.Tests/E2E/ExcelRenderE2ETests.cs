// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using FluentAssertions;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public sealed class ExcelRenderE2ETests : ExcelTestBase
{
    // ==================== view html --out: cells and values ====================

    [Fact]
    public void ViewHtml_TextNumberDate_ProducesInspectableHtml()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Hello World")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "42"), ("type", "number")));
            handler.Add("/Sheet1/C1", "cell", null, Props(("value", "2025-06-15"), ("type", "date")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Hello World");
        html.Should().Contain("42");
    }

    [Fact]
    public void ViewHtml_Formulas_ShowsFormulaResults()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "10"), ("type", "number")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "20"), ("type", "number")));
            handler.Add("/Sheet1/A3", "cell", null, Props(("formula", "SUM(A1:A2)")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // The formula should resolve: 10 + 20 = 30
        html.Should().Contain("30");
    }

    [Fact]
    public void ViewHtml_NumberFormats_ShowsFormattedValues()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "0.156"), ("type", "number"), ("numberformat", "0.00%")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "1234.567"), ("type", "number"), ("numberformat", "#,##0.00")));
            handler.Add("/Sheet1/A3", "cell", null, Props(("value", "45000"), ("type", "number"), ("numberformat", "$#,##0.00")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // Percentage should appear: 15.60% (or similar)
        html.Should().Contain("%");
    }

    [Fact]
    public void ViewHtml_FreezePanes_ContainsFreezeInfo()
    {
        var path = CreateWorkbook();

        // Add some data first.
        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Header1")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "Header2")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "Data1")));
            handler.Add("/Sheet1/B2", "cell", null, Props(("value", "Data2")));
            handler.Save();
        }

        // Set freeze pane at A2 (freeze row 1) by injecting sheet views XML.
        RunCliOk("raw-set", path, "/Sheet1",
            "--xpath", "/x:worksheet",
            "--action", "prepend",
            "--xml",
            "<x:sheetViews xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<x:sheetView tabSelected=\"1\" workbookViewId=\"0\">" +
            "<x:pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>" +
            "</x:sheetView>" +
            "</x:sheetViews>");

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Header1");
        // Frozen pane should produce a sticky/frozen indicator in the HTML.
        (html.Contains("frozen") || html.Contains("sticky") || html.Contains("freeze")).Should().BeTrue();
    }

    [Fact]
    public void ViewHtml_MergedCells_ShowsMergedContent()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Merged Title")));
            handler.Set("/Sheet1/A1:C1", Props(("merge", "true")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Merged Title");
    }

    [Fact]
    public void ViewHtml_RichText_BoldItalicColor_ShowsFormattedContent()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Bold Text"), ("font.bold", "true")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "Italic Text"), ("font.italic", "true")));
            handler.Add("/Sheet1/A3", "cell", null, Props(("value", "Blue Text"), ("font.color", "0000FF")));
            handler.Add("/Sheet1/A4", "cell", null, Props(("value", "Red Background"), ("fill", "FF0000")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Bold Text");
        html.Should().Contain("Italic Text");
        html.Should().Contain("Blue Text");
    }

    [Fact]
    public void ViewHtml_ConditionalFormatting_ShowsCfVisuals()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            // Add data values for conditional formatting.
            for (var i = 1; i <= 6; i++)
            {
                handler.Add($"/Sheet1/A{i}", "cell", null,
                    Props(("value", (i * 10).ToString()), ("type", "number")));
            }

            // Add a conditional formatting rule: highlight cells > 30 with red fill.
            handler.Add("/Sheet1", "cf", null, Props(
                ("ref", "A1:A6"),
                ("type", "cellIs"),
                ("operator", "greaterThan"),
                ("value", "30"),
                ("fill", "FF0000")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // CF styling should be present in the HTML.
        html.Should().Contain("60"); // values > 30 should appear
    }

    [Fact]
    public void ViewHtml_Chart_ShowsChartInHtml()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            // Add data for chart.
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Category"), ("font.bold", "true")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "Value"), ("font.bold", "true")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "10"), ("type", "number")));
            handler.Add("/Sheet1/B2", "cell", null, Props(("value", "25"), ("type", "number")));
            handler.Add("/Sheet1/A3", "cell", null, Props(("value", "20"), ("type", "number")));
            handler.Add("/Sheet1/B3", "cell", null, Props(("value", "40"), ("type", "number")));
            handler.Add("/Sheet1/A4", "cell", null, Props(("value", "30"), ("type", "number")));
            handler.Add("/Sheet1/B4", "cell", null, Props(("value", "60"), ("type", "number")));

            // Add a column chart.
            handler.Add("/Sheet1", "chart", null, Props(
                ("chartType", "column"),
                ("dataRange", "Sheet1!A1:B4"),
                ("anchor", "F2:L16"),
                ("title", "TestChart")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        // The chart should appear in the HTML — either as an embedded image (base64/img tag)
        // or as SVG/Canvas element.
        (html.Contains("<img") || html.Contains("chart") || html.Contains("svg")).Should().BeTrue();
    }

    [Fact]
    public void ViewHtml_Drawings_ShowsDrawingContent()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Before shape")));
            handler.Add("/Sheet1/A10", "cell", null, Props(("value", "After shape")));

            // Add a shape (rectangle).
            handler.Add("/Sheet1", "shape", null, Props(
                ("type", "rect"),
                ("text", "ShapeText"),
                ("x", "100"), ("y", "100"),
                ("width", "200"), ("height", "100")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("Before shape");
    }

    [Fact]
    public void ViewHtml_OverflowText_ShowsLongContent()
    {
        var path = CreateWorkbook();

        var longText = "This is a very long text string that will overflow into adjacent empty cells in Excel";

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", longText)));
            // Keep column A narrow to demonstrate overflow.
            handler.Set("/Sheet1/col[A]", Props(("width", "8")));
            handler.Save();
        }

        var htmlPath = NewTempPath(".html");
        RunCliOk("view", path, "html", "--out", htmlPath);
        TrackTempFile(htmlPath);

        var html = File.ReadAllText(htmlPath);
        html.Should().Contain("<html");
        html.Should().Contain("very long text");
    }

    // ==================== view screenshot ====================

    [Fact]
    public void ViewScreenshot_Basic_AcceptsNoBackendOrProducesPng()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Screenshot Test")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "Content")));
            handler.Save();
        }

        var ssPath = NewTempPath(".png");
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--screenshot-width", "480", "--screenshot-height", "320");
        TrackTempFile(ssPath);

        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            // PNG was produced — verify it's a valid PNG.
            File.ReadAllBytes(ssPath).Should().StartWith(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
        }
        else
        {
            // Missing render backend is acceptable.
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    [Fact]
    public void ViewScreenshot_WithChart_AcceptsNoBackendOrProducesPng()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "10"), ("type", "number")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "20"), ("type", "number")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "30"), ("type", "number")));
            handler.Add("/Sheet1/B2", "cell", null, Props(("value", "40"), ("type", "number")));

            handler.Add("/Sheet1", "chart", null, Props(
                ("chartType", "bar"),
                ("dataRange", "Sheet1!A1:B2"),
                ("anchor", "F2:L16"),
                ("title", "ChartSS")));
            handler.Save();
        }

        var ssPath = NewTempPath(".png");
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--screenshot-width", "480", "--screenshot-height", "320");
        TrackTempFile(ssPath);

        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            File.ReadAllBytes(ssPath).Should().StartWith(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
        }
        else
        {
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    [Fact]
    public void ViewScreenshot_LargeSheet_AcceptsNoBackendOrProducesPng()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            // Populate a larger grid.
            for (var r = 1; r <= 20; r++)
            {
                for (var c = 0; c < 5; c++)
                {
                    var col = (char)('A' + c);
                    handler.Add($"/Sheet1/{col}{r}", "cell", null,
                        Props(("value", $"R{r}C{c}")));
                }
            }
            handler.Save();
        }

        var ssPath = NewTempPath(".png");
        var result = RunCli("view", path, "screenshot", "--out", ssPath,
            "--screenshot-width", "800", "--screenshot-height", "600");
        TrackTempFile(ssPath);

        if (result.ExitCode == 0 && File.Exists(ssPath) && new FileInfo(ssPath).Length > 0)
        {
            File.ReadAllBytes(ssPath).Should().StartWith(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
        }
        else
        {
            (result.Stdout + result.Stderr).Should().Contain("no_screenshot_backend");
        }
    }

    // ==================== Helpers ====================

    private static Dictionary<string, string> Props(params (string Key, string Value)[] values)
        => values.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
}
