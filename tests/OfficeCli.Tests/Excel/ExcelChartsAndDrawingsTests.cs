// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelChartsAndDrawingsTests : ExcelTestBase
{
    [Fact]
    public void ChartsSeriesAxesPicturesShapesOleSparklinesAndBreaks_RoundTrip()
    {
        var path = CreateWorkbook();
        var imagePath = CreateTinyPng();
        var olePath = CreateOlePayload();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Revenue", "Cost");
            AddRow(handler, 2, "Jan", "100", "60");
            AddRow(handler, 3, "Feb", "120", "70");
            AddRow(handler, 4, "Mar", "90", "55");

            handler.Add("/Sheet1", "chart", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["type"] = "column",
                ["title"] = "Revenue",
                ["dataRange"] = "Sheet1!A1:C4",
                ["anchor"] = "E2:L16",
                ["name"] = "Revenue Chart",
                ["axisTitle"] = "Amount"
            });
            handler.Add("/Sheet1/chart[1]", "chart-series", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "Forecast",
                ["values"] = "130,140,150",
                ["categories"] = "Jan,Feb,Mar",
                ["color"] = "70AD47"
            });
            handler.Set("/Sheet1/chart[1]/axis[@role=value]", new(StringComparer.OrdinalIgnoreCase)
            {
                ["title"] = "USD",
                ["visible"] = "true"
            });
            handler.Add("/Sheet1", "picture", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["src"] = imagePath,
                ["alt"] = "Logo",
                ["title"] = "Logo title",
                ["anchor"] = "A6:C10",
                ["rotation"] = "15",
                ["flip"] = "h",
                ["hyperlink"] = "https://officecli.ai"
            });
            handler.Add("/Sheet1", "shape", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["text"] = "Callout",
                ["name"] = "Callout Shape",
                ["anchor"] = "D6:F10",
                ["geometry"] = "roundRect",
                ["fill"] = "4472C4",
                ["font.color"] = "FFFFFF",
                ["font.bold"] = "true"
            });
            handler.Add("/Sheet1", "ole", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["src"] = olePath,
                ["progId"] = "Package",
                ["anchor"] = "G6:I10"
            });
            handler.Add("/Sheet1/D2", "sparkline", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["dataRange"] = "B2:C2",
                ["type"] = "column",
                ["color"] = "4472C4",
                ["negativeColor"] = "FF0000",
                ["markers"] = "true",
                ["highPoint"] = "true",
                ["lowPoint"] = "true",
                ["lineWeight"] = "1.5",
                ["displayEmptyCellsAs"] = "zero"
            });
            handler.Add("/Sheet1", "rowbreak", null, new(StringComparer.OrdinalIgnoreCase) { ["row"] = "20" });
            handler.Add("/Sheet1", "pagebreak", null, new(StringComparer.OrdinalIgnoreCase) { ["col"] = "F" });

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var chart = Query(readOnly, "chart").Should().ContainSingle().Subject;
        chart.Path.Should().Be("/Sheet1/chart[1]");
        var series = ReadNode(readOnly, "/Sheet1/chart[1]/series[1]");
        series.Type.Should().Be("series");
        ReadNode(readOnly, "/Sheet1/chart[1]/axis[@role=value]").Format.Should().Contain("role", "value");

        var picture = Query(readOnly, "picture").Should().ContainSingle().Subject;
        picture.Path.Should().Be("/Sheet1/picture[1]");
        picture.Format.Should().Contain("alt", "Logo");

        var shape = Query(readOnly, "shape").Should().ContainSingle().Subject;
        shape.Path.Should().Be("/Sheet1/shape[1]");
        shape.Text.Should().Be("Callout");
        shape.Format.Should().Contain("name", "Callout Shape");
        shape.Format.Should().Contain("geometry", "roundRect");

        var ole = Query(readOnly, "ole").Should().ContainSingle().Subject;
        ole.Path.Should().Be("/Sheet1/ole[1]");
        ole.Format.Should().Contain("progId", "Package");
        ole.Format.Should().ContainKey("fileSize");

        var sparkline = Query(readOnly, "sparkline").Should().ContainSingle().Subject;
        sparkline.Path.Should().Be("/Sheet1/sparkline[1]");
        sparkline.Format.Should().Contain("type", "column");
        sparkline.Format.Should().Contain("dataRange", "B2:C2");
        sparkline.Format.Should().Contain("location", "D2");
        sparkline.Format.Should().Contain("color", "#4472C4");

        var rowBreak = ReadNode(readOnly, "/Sheet1/rowbreak[1]");
        rowBreak.Format.Should().Contain("row", 20U);
        rowBreak.Format.Should().Contain("manual", true);
        var colBreak = ReadNode(readOnly, "/Sheet1/colbreak[1]");
        colBreak.Format.Should().Contain("col", 6);
        colBreak.Format.Should().Contain("manual", true);
        readOnly.Validate().Should().BeEmpty();
    }

    private string CreateTinyPng()
    {
        var path = TrackTempFile(Path.Combine(Path.GetTempPath(), $"officecli_excel_{Guid.NewGuid():N}.png"));
        File.WriteAllBytes(path, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
        return path;
    }

    private string CreateOlePayload()
    {
        var path = TrackTempFile(Path.Combine(Path.GetTempPath(), $"officecli_excel_{Guid.NewGuid():N}.txt"));
        File.WriteAllText(path, "embedded payload");
        return path;
    }

    private static void AddRow(OfficeCli.Handlers.ExcelHandler handler, int row, params string[] values)
    {
        for (var i = 0; i < values.Length; i++)
        {
            var column = (char)('A' + i);
            handler.Add($"/Sheet1/{column}{row}", "cell", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["value"] = values[i]
            });
        }
    }
}
