// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class ExcelChartsDrawingsIntegrationTests : ExcelTestBase
{
    // ── Standard charts ────────────────────────────────────────────────

    [Fact]
    public void Chart_Column_AddGetSetRemoveWithSaveReopen_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Sales");
            AddRow(handler, 2, "Jan", "100");
            AddRow(handler, 3, "Feb", "120");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "column"),
                ("dataRange", "Sheet1!A1:B3"),
                ("anchor", "F2:L16"),
                ("title", "Sales Chart"),
                ("xlabel", "Month"),
                ("ylabel", "Amount"),
                ("legend", "true")
            ));
            handler.Set("/Sheet1/chart[1]", Props(
                ("title", "Updated Sales"),
                ("legend", "false")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var chart = Query(readOnly, "chart").Should().ContainSingle(n => n.Path == "/Sheet1/chart[1]").Subject;
        chart.Format.Should().ContainKey("chartType");
        var getChart = ReadNode(readOnly, "/Sheet1/chart[1]");
        getChart.Should().NotBeNull();
        readOnly.Validate().Should().BeEmpty();

        // Remove and verify
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/chart[1]", null);
            handler.Save();
        }
        using var finalReadOnly = OpenReadOnly(path);
        Query(finalReadOnly, "chart").Should().BeEmpty();
        finalReadOnly.Validate().Should().BeEmpty();
    }

    [Theory]
    [InlineData("bar")]
    [InlineData("line")]
    [InlineData("pie")]
    [InlineData("area")]
    public void Chart_MultipleTypes_AddAndReadback(string chartType)
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales");
            AddRow(handler, 2, "West", "1200");
            AddRow(handler, 3, "East", "800");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", chartType),
                ("dataRange", "Sheet1!A1:B3"),
                ("anchor", "F2:L16"),
                ("title", chartType + " Chart")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle(n => n.Path == "/Sheet1/chart[1]");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Chart_ComboType_AddAndReadback()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales", "Cost");
            AddRow(handler, 2, "West", "1200", "600");
            AddRow(handler, 3, "East", "800", "400");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "combo"),
                ("dataRange", "Sheet1!A1:C3"),
                ("anchor", "F2:L16"),
                ("title", "Combo Chart")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle(n => n.Path == "/Sheet1/chart[1]");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Extended charts ─────────────────────────────────────────────────

    [Theory]
    [InlineData("waterfall")]
    [InlineData("histogram")]
    [InlineData("boxWhisker")]
    public void Chart_ExtendedTypes_AddAndReadback(string chartType)
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Category", "Value");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "20");
            AddRow(handler, 4, "C", "15");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", chartType),
                ("dataRange", "Sheet1!A1:B4"),
                ("anchor", "F2:L16"),
                ("title", chartType + " Chart")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle(n => n.Path == "/Sheet1/chart[1]");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Chart series ────────────────────────────────────────────────────

    [Fact]
    public void ChartSeries_AddAndReadback()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Sales");
            AddRow(handler, 2, "Jan", "100");
            AddRow(handler, 3, "Feb", "120");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "column"),
                ("dataRange", "Sheet1!A1:B3"),
                ("anchor", "F2:L16"),
                ("title", "Chart with Series")
            ));
            handler.Add("/Sheet1/chart[1]", "chart-series", null, Props(
                ("name", "Forecast"),
                ("values", "130,140"),
                ("categories", "Mar,Apr")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle();
        var series = ReadNode(readOnly, "/Sheet1/chart[1]/series[1]");
        series.Type.Should().Be("series");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Chart properties ─────────────────────────────────────────────────

    [Fact]
    public void ChartProperties_LayoutTitleXlabelYlabelLegend_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Quarter", "Revenue");
            AddRow(handler, 2, "Q1", "500");
            AddRow(handler, 3, "Q2", "600");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "bar"),
                ("dataRange", "Sheet1!A1:B3"),
                ("anchor", "F2:L16"),
                ("title", "Quarterly Revenue"),
                ("xlabel", "Quarter"),
                ("ylabel", "USD"),
                ("legend", "true"),
                ("layout", "1")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle();
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Chart_TrendlineAndErrorBars_SetAndSave()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Sales");
            AddRow(handler, 2, "Jan", "100");
            AddRow(handler, 3, "Feb", "120");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "column"),
                ("dataRange", "Sheet1!A1:B3"),
                ("anchor", "F2:L16"),
                ("title", "Trend Chart")
            ));
            handler.Set("/Sheet1/chart[1]", Props(
                ("trendline", "linear"),
                ("errorBars", "standardError")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle();
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Pictures ────────────────────────────────────────────────────────

    [Fact]
    public void Picture_AddGetSaveRemove_RoundTrip()
    {
        var path = CreateWorkbook();
        var png = CreateTinyPng(Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png"));

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Data");
            handler.Add("/Sheet1", "picture", null, Props(
                ("src", png),
                ("alt", "Logo"),
                ("anchor", "A5:C9"),
                ("hyperlink", "https://example.com")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "picture").Should().ContainSingle(n => n.Path == "/Sheet1/picture[1]");
        var pic = ReadNode(readOnly, "/Sheet1/picture[1]");
        pic.Format.Should().Contain("alt", "Logo");
        readOnly.Validate().Should().BeEmpty();

        // Remove picture
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/picture[1]", null);
            handler.Save();
        }
        using var finalReadOnly = OpenReadOnly(path);
        Query(finalReadOnly, "picture").Should().BeEmpty();
        finalReadOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Picture_GetMetadata_ReturnsExpectedKeys()
    {
        var path = CreateWorkbook();
        var png = CreateTinyPng(Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png"));

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "picture", null, Props(
                ("src", png),
                ("alt", "MetadataTest"),
                ("anchor", "B2:D6")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var pic = ReadNode(readOnly, "/Sheet1/picture[1]");
        pic.Format.Should().Contain("alt", "MetadataTest");
        pic.Format.Should().ContainKey("name");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Shapes ──────────────────────────────────────────────────────────

    [Fact]
    public void Shape_AddSetRemoveWithFormatting_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "shape", null, Props(
                ("text", "Hello World"),
                ("anchor", "D5:F9"),
                ("geometry", "rect"),
                ("font.bold", "true"),
                ("font.size", "14"),
                ("font.color", "FF0000"),
                ("fill", "C6EFCE"),
                ("line", "000000"),
                ("align", "center"),
                ("valign", "middle"),
                ("margins", "91440,45720,91440,45720")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "shape").Should().ContainSingle(n => n.Path == "/Sheet1/shape[1]");
        var shape = ReadNode(readOnly, "/Sheet1/shape[1]");
        shape.Text.Should().Be("Hello World");
        shape.Format.Should().Contain("geometry", "rect");
        readOnly.Validate().Should().BeEmpty();

        // Remove shape
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/shape[1]", null);
            handler.Save();
        }
        using var finalReadOnly = OpenReadOnly(path);
        Query(finalReadOnly, "shape").Should().BeEmpty();
        finalReadOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Shape_RoundedRectGeometry()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "shape", null, Props(
                ("text", "Rounded"),
                ("anchor", "A1:C5"),
                ("geometry", "roundedRect")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var shape = Query(readOnly, "shape").Should().ContainSingle().Subject;
        shape.Text.Should().Be("Rounded");
        shape.Format.Should().ContainKey("geometry");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── OLE objects ─────────────────────────────────────────────────────

    [Fact]
    public void Ole_AddGetRemoveWithPayload_RoundTrip()
    {
        var path = CreateWorkbook();
        var ole = CreateOlePayload(Path.Combine(Path.GetTempPath(), $"ole_{Guid.NewGuid():N}.txt"));

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "ole", null, Props(
                ("src", ole),
                ("progId", "Package"),
                ("anchor", "G5:I9")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "ole").Should().ContainSingle(n => n.Path == "/Sheet1/ole[1]");
        var oleNode = ReadNode(readOnly, "/Sheet1/ole[1]");
        oleNode.Format.Should().Contain("progId", "Package");
        oleNode.Format.Should().ContainKey("fileSize");
        readOnly.Validate().Should().BeEmpty();

        // Remove OLE
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/ole[1]", null);
            handler.Save();
        }
        using var finalReadOnly = OpenReadOnly(path);
        Query(finalReadOnly, "ole").Should().BeEmpty();
        finalReadOnly.Validate().Should().BeEmpty();
    }

    // ── Sparklines ──────────────────────────────────────────────────────

    [Fact]
    public void Sparkline_LineAndColumn_AddAndReadback()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Sales", "Cost");
            AddRow(handler, 2, "Jan", "10", "7");
            AddRow(handler, 3, "Feb", "12", "8");
            handler.Add("/Sheet1/D2", "sparkline", null, Props(
                ("dataRange", "B2:C2"),
                ("type", "line")
            ));
            handler.Add("/Sheet1/D3", "sparkline", null, Props(
                ("dataRange", "B3:C3"),
                ("type", "column")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "sparkline").Should().HaveCount(2);
        var sparklines = Query(readOnly, "sparkline");
        sparklines.Select(n => n.Format["type"]!.ToString()).Should().Contain(new[] { "line", "column" });
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Sparkline_QueryAndReadback()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Sales");
            AddRow(handler, 2, "Jan", "10");
            handler.Add("/Sheet1/C2", "sparkline", null, Props(
                ("dataRange", "B2"),
                ("type", "column")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sparkline = Query(readOnly, "sparkline").Should().ContainSingle().Subject;
        sparkline.Format.Should().Contain("location", "C2");
        sparkline.Format.Should().Contain("type", "column");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Query all drawings ─────────────────────────────────────────────

    [Fact]
    public void QueryAllDrawings_ReturnsChartPictureShapeSparklineOle()
    {
        var path = CreateWorkbook();
        var png = CreateTinyPng(Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png"));
        var ole = CreateOlePayload(Path.Combine(Path.GetTempPath(), $"ole_{Guid.NewGuid():N}.txt"));

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Sales", "Cost");
            AddRow(handler, 2, "Jan", "10", "7");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "column"),
                ("dataRange", "Sheet1!A1:B2"),
                ("anchor", "F2:L16"),
                ("title", "Sales")
            ));
            handler.Add("/Sheet1", "picture", null, Props(
                ("src", png),
                ("alt", "Logo"),
                ("anchor", "A5:C9")
            ));
            handler.Add("/Sheet1", "shape", null, Props(
                ("text", "Note"),
                ("anchor", "D5:F9"),
                ("geometry", "rect")
            ));
            handler.Add("/Sheet1/D2", "sparkline", null, Props(
                ("dataRange", "B2:C2"),
                ("type", "line")
            ));
            handler.Add("/Sheet1", "ole", null, Props(
                ("src", ole),
                ("progId", "Package"),
                ("anchor", "G5:I9")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle();
        Query(readOnly, "picture").Should().ContainSingle();
        Query(readOnly, "shape").Should().ContainSingle();
        Query(readOnly, "sparkline").Should().ContainSingle();
        Query(readOnly, "ole").Should().ContainSingle();
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Raw chart/drawing access ───────────────────────────────────────

    [Fact]
    public void RawAccess_ChartAndDrawingParts_ContainExpectedXml()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Sales");
            AddRow(handler, 2, "Jan", "100");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "column"),
                ("dataRange", "Sheet1!A1:B2"),
                ("anchor", "F2:L16"),
                ("title", "RawTest")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var chartRaw = readOnly.Raw("/Sheet1/chart[1]");
        chartRaw.Should().NotBeNullOrWhiteSpace();
        var drawingRaw = readOnly.Raw("/Sheet1/drawing");
        drawingRaw.Should().Contain("wsDr");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Helpers ─────────────────────────────────────────────────────────

    private static Dictionary<string, string> Props(params (string Key, string Value)[] values)
        => values.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

    private static void AddRow(ExcelHandler handler, int row, params string[] values)
    {
        for (var i = 0; i < values.Length; i++)
        {
            var column = (char)('A' + i);
            handler.Add($"/Sheet1/{column}{row}", "cell", null, Props(("value", values[i])));
        }
    }
}
