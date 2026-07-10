// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class ExcelIntegrationTests : ExcelTestBase
{
    [Fact]
    public void WorkbookAuthoring_ComposesTablesChartsCfValidationAndFormulas()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales", "Status");
            AddRow(handler, 2, "West", "1200", "Open");
            AddRow(handler, 3, "East", "800", "Closed");
            handler.Add("/Sheet1/D1", "cell", null, Props(("formula", "SUM(B2:B3)")));
            handler.Add("/Sheet1", "table", null, Props(("ref", "A1:C3"), ("name", "SalesTable"), ("columns", "Region,Sales,Status")));
            handler.Add("/", "namedrange", null, Props(("name", "TotalSales"), ("ref", "Sheet1!$D$1")));
            handler.Add("/Sheet1", "validation", null, Props(("ref", "C2:C3"), ("type", "list"), ("formula1", "Open,Closed")));
            handler.Add("/Sheet1", "cellis", null, Props(("range", "B2:B3"), ("operator", "greaterThan"), ("formula", "1000"), ("fill", "C6EFCE")));
            handler.Add("/Sheet1", "chart", null, Props(("type", "column"), ("dataRange", "Sheet1!A1:B3"), ("anchor", "F2:L16"), ("title", "Sales")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/D1").Format.Should().Contain("formula", "SUM(B2:B3)");
        Query(readOnly, "listobject").Should().ContainSingle(n => n.Path == "/Sheet1/table[1]");
        Query(readOnly, "namedrange").Should().ContainSingle(n => n.Format.ContainsValue("TotalSales"));
        Query(readOnly, "validation").Should().ContainSingle();
        ReadNode(readOnly, "/Sheet1/cf[1]").Format.Should().Contain("type", "cellIs");
        Query(readOnly, "chart").Should().ContainSingle(n => n.Path == "/Sheet1/chart[1]");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookMutation_RewritesReferencesAcrossSheetColumnMoves()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1", Props(("printArea", "A1:C20")));
            handler.Add("/Sheet1", "row", null, Props(("cols", "3")));
            handler.Add("/Sheet1", "col", null, Props(("name", "B"), ("width", "18")));
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "A")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "B")));
            handler.Add("/Sheet1/C1", "cell", null, Props(("formula", "SUM(A1:B1)")));
            handler.Add("/", "namedrange", null, Props(("name", "MovedSource"), ("ref", "Sheet1!$A$1:$C$1")));
            handler.Move("/Sheet1/col[A]", null, OfficeCli.Core.InsertPosition.AfterElement("/Sheet1/col[C]"));
            handler.Remove("/Sheet1/C1", null);
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1").Format.Should().Contain("printArea", "A1:D20");
        ReadNode(readOnly, "/Sheet1/col[A]").Format.Should().Contain("width", 18D);
        Query(readOnly, "namedrange").Should().ContainSingle(n => n.Format["name"]!.Equals("MovedSource"));
        Query(readOnly, "cell").Select(n => n.Path).Should().NotContain("/Sheet1/D1");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookDrawingRoundTrip_PreservesPicturesShapesOleAndSparklines()
    {
        var path = CreateWorkbook();
        var image = CreateTinyPng();
        var ole = CreateOlePayload();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Month", "Sales", "Cost");
            AddRow(handler, 2, "Jan", "10", "7");
            AddRow(handler, 3, "Feb", "12", "8");
            handler.Add("/Sheet1", "picture", null, Props(("src", image), ("alt", "Logo"), ("anchor", "A5:C9")));
            handler.Add("/Sheet1", "shape", null, Props(("text", "Note"), ("anchor", "D5:F9"), ("geometry", "rect")));
            handler.Add("/Sheet1", "ole", null, Props(("src", ole), ("progId", "Package"), ("anchor", "G5:I9")));
            handler.Add("/Sheet1/D2", "sparkline", null, Props(("dataRange", "B2:C2"), ("type", "line")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "picture").Should().ContainSingle();
        Query(readOnly, "shape").Should().ContainSingle(n => n.Text == "Note");
        Query(readOnly, "ole").Should().ContainSingle(n => n.Format.ContainsKey("fileSize"));
        Query(readOnly, "sparkline").Should().ContainSingle(n => n.Format["location"]!.Equals("D2"));
        readOnly.Raw("/Sheet1/drawing").Should().Contain("wsDr");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookDumpBatchReplay_RecreatesSemantics()
    {
        var source = CreateWorkbook();
        var target = CreateWorkbook();

        using (var handler = OpenEditable(source))
        {
            AddRow(handler, 1, "Item", "Qty");
            AddRow(handler, 2, "Pen", "2");
            handler.Add("/Sheet1/B3", "cell", null, Props(("formula", "SUM(B2:B2)")));
            handler.Add("/", "namedrange", null, Props(("name", "QtyTotal"), ("ref", "Sheet1!$B$3")));
            handler.Save();
        }

        List<BatchItem> items;
        using (var readOnly = OpenReadOnly(source))
            (items, _) = ExcelBatchEmitter.EmitExcel(readOnly);

        using (var replay = OpenEditable(target))
        {
            var results = CommandBuilder.RunNonResidentBatch(replay, items, stopOnError: true, json: false);
            results.Should().OnlyContain(r => r.Success);
            replay.Save();
        }

        using var targetReadOnly = OpenReadOnly(target);
        ReadNode(targetReadOnly, "/Sheet1/A2").Text.Should().Be("Pen");
        ReadNode(targetReadOnly, "/Sheet1/B3").Format.Should().Contain("formula", "SUM(B2:B2)");
        Query(targetReadOnly, "namedrange").Should().ContainSingle(n => n.Format["name"]!.Equals("QtyTotal"));
        targetReadOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookImportViewAndQuery_ComposesBulkDataWithRenderReadback()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            var result = handler.Import("/Sheet1", "Region,Sales,Status\nWest,1200,Open\nEast,800,Closed", ',', hasHeader: true, startCell: "A1");
            result.Should().Contain("Imported");
            handler.Set("/Sheet1/col[A]", Props(("width", "18")));
            handler.Set("/Sheet1/col[B]", Props(("width", "12")));
            handler.Set("/Sheet1/B4", Props(("formula", "SUM(B2:B3)"), ("numberformat", "#,##0")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        readOnly.ViewAsText(cols: new HashSet<string> { "A", "B" })
            .Should().Contain("A2=West").And.Contain("B4=2000");
        var html = readOnly.ViewAsHtml();
        html.Should().Contain("<html").And.Contain("West").And.NotContain("###");
        Query(readOnly, "cell[value>=1000]").Select(n => n.Path).Should().Contain("/Sheet1/B2");
        Query(readOnly, "cell[text~=Closed]").Select(n => n.Path).Should().Contain("/Sheet1/C3");
        readOnly.Raw("/Sheet1", startRow: 1, endRow: 2, cols: new HashSet<string> { "A", "B" })
            .Should().Contain("West").And.Contain("1200");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookAnalytics_ComposesPivotSlicerAndChartOnSharedData()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Category", "Sales", "Qty");
            AddRow(handler, 2, "West", "Hardware", "1200", "2");
            AddRow(handler, 3, "East", "Software", "800", "5");
            AddRow(handler, 4, "West", "Software", "1500", "3");
            AddRow(handler, 5, "North", "Hardware", "600", "1");
            handler.Add("/Sheet1", "chart", null, Props(("type", "bar"), ("dataRange", "Sheet1!A1:C5"), ("anchor", "J2:P16"), ("title", "Sales by Region")));
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:D5"),
                ("position", "F1"),
                ("name", "SalesPivot"),
                ("rows", "Region"),
                ("cols", "Category"),
                ("values", "Sales:sum,Qty:count"),
                ("style", "PivotStyleMedium9")));
            handler.Add("/Sheet1", "slicer", null, Props(
                ("pivotTable", "/Sheet1/pivottable[1]"),
                ("field", "Region"),
                ("name", "RegionSlicer"),
                ("caption", "Filter Region"),
                ("columnCount", "2"),
                ("style", "SlicerStyleLight1")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle(n => n.Path == "/Sheet1/chart[1]");
        Query(readOnly, "pivottable").Should().ContainSingle(n =>
            n.Format["name"]!.Equals("SalesPivot") && n.Format.ContainsKey("dataField1"));
        Query(readOnly, "slicer").Should().ContainSingle(n =>
            n.Format["name"]!.Equals("RegionSlicer") && n.Format.ContainsKey("pivotCacheId"));
        readOnly.Raw("/Sheet1/drawing").Should().Contain("graphicFrame");
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
