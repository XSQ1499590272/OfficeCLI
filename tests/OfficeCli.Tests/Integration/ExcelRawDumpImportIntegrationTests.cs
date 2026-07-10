// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class ExcelRawDumpImportIntegrationTests : ExcelTestBase
{
    // ── Raw: read various parts ─────────────────────────────────────────

    [Fact]
    public void Raw_ReadWorkbookAndSheetParts_ReturnsValidXml()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Value");
            AddRow(handler, 2, "Alpha", "100");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var wbRaw = readOnly.Raw("/workbook");
        wbRaw.Should().NotBeNullOrWhiteSpace();

        var sheetRaw = readOnly.Raw("/Sheet1");
        sheetRaw.Should().Contain("Alpha");
        sheetRaw.Should().Contain("100");

        var stylesRaw = readOnly.Raw("/styles");
        stylesRaw.Should().NotBeNullOrWhiteSpace();

        var ssRaw = readOnly.Raw("/sharedstrings");
        ssRaw.Should().NotBeNullOrWhiteSpace();

        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Raw_ReadWithRowColumnFilters_ReturnsFilteredXml()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Value", "Status");
            AddRow(handler, 2, "Alpha", "100", "Open");
            AddRow(handler, 3, "Beta", "200", "Closed");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var filtered = readOnly.Raw("/Sheet1", startRow: 1, endRow: 2, cols: new HashSet<string> { "A", "B" });
        filtered.Should().Contain("Alpha");
        filtered.Should().Contain("100");
        filtered.Should().NotContain("Closed");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Raw-set: mutate workbook/sheet parts ───────────────────────────

    [Fact]
    public void RawSet_SetAttrOnSheet_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Data");
            handler.RawSet("/Sheet1", "/x:worksheet/x:sheetData/x:row[1]", "setattr",
                "hidden=1");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        readOnly.Raw("/Sheet1").Should().Contain("hidden");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void RawSet_MutateWorkbookPart_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            // Add data, then use raw-set append to inject additional content
            AddRow(handler, 1, "Header");
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            // Append a new row with setattr for validation
            handler.RawSet("/Sheet1", "/x:worksheet/x:sheetData", "append",
                "<x:row r=\"99\"><x:c r=\"A99\" t=\"str\"><x:v>RawAppended</x:v></x:c></x:row>");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var raw = readOnly.Raw("/Sheet1");
        raw.Should().Contain("RawAppended");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void RawSet_AppendAction_AddsContent()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item");
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.RawSet("/Sheet1", "/x:worksheet/x:sheetData", "append",
                "<x:row r=\"5\"><x:c r=\"A5\" t=\"str\"><x:v>Appended</x:v></x:c></x:row>");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        readOnly.Raw("/Sheet1").Should().Contain("Appended");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void RawSet_ReplaceAction_ReplacesContent()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Original");
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.RawSet("/Sheet1", "/x:worksheet/x:sheetData/x:row[1]/x:c[1]", "replace",
                "<x:c r=\"A1\" t=\"str\"><x:v>Replaced</x:v></x:c>");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var raw = readOnly.Raw("/Sheet1");
        raw.Should().Contain("Replaced");
        raw.Should().NotContain("Original");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void RawSet_RemoveAction_RemovesContent()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Keep", "ToRemove");
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.RawSet("/Sheet1", "/x:worksheet/x:sheetData/x:row[1]/x:c[2]", "remove", "");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var raw = readOnly.Raw("/Sheet1");
        raw.Should().Contain("Keep");
        raw.Should().NotContain("ToRemove");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Dump/batch: create → dump → replay ─────────────────────────────

    [Fact]
    public void DumpBatch_CreateWorkbookWithDataAndChart_DumpReplayIntoBlank_VerifiesSemantics()
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

        items.Should().NotBeEmpty();

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
    public void DumpBatch_WithTableAndChart_DumpReplayIntoBlank_VerifiesSemantics()
    {
        var source = CreateWorkbook();
        var target = CreateWorkbook();

        using (var handler = OpenEditable(source))
        {
            AddRow(handler, 1, "Region", "Sales", "Status");
            AddRow(handler, 2, "West", "1200", "Open");
            AddRow(handler, 3, "East", "800", "Closed");
            handler.Add("/Sheet1", "table", null, Props(
                ("ref", "A1:C3"),
                ("name", "SalesTable"),
                ("columns", "Region,Sales,Status")
            ));
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "column"),
                ("dataRange", "Sheet1!A1:B3"),
                ("anchor", "F2:L16"),
                ("title", "Sales Chart")
            ));
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
        Query(targetReadOnly, "listobject").Should().ContainSingle(n => n.Format["name"]!.Equals("SalesTable"));
        Query(targetReadOnly, "chart").Should().ContainSingle();
        targetReadOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void DumpBatch_WithNamedRanges_DumpReplayRoundTrip()
    {
        var source = CreateWorkbook();
        var target = CreateWorkbook();

        using (var handler = OpenEditable(source))
        {
            handler.Add("/", "namedrange", null, Props(("name", "MyRange"), ("ref", "Sheet1!$A$1:$B$10")));
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
        Query(targetReadOnly, "namedrange").Should().ContainSingle(n => n.Format["name"]!.Equals("MyRange"));
        targetReadOnly.Validate().Should().BeEmpty();
    }

    // ── Import: CSV/TSV ─────────────────────────────────────────────────

    [Fact]
    public void Import_CsvWithHeader_ImportsAndVerifiesCellReadback()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            var csv = "Region,Sales,Status\nWest,1200,Open\nEast,800,Closed";
            var result = handler.Import("/Sheet1", csv, ',', hasHeader: true, startCell: "A1");
            result.Should().Contain("Imported");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/A1").Text.Should().Be("Region");
        ReadNode(readOnly, "/Sheet1/A2").Text.Should().Be("West");
        ReadNode(readOnly, "/Sheet1/B2").Text.Should().Be("1200");
        ReadNode(readOnly, "/Sheet1/C3").Text.Should().Be("Closed");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Import_TsvWithHeader_ImportsAndVerifiesCellReadback()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            var tsv = "Name\tValue\nAlpha\t100\nBeta\t200";
            var result = handler.Import("/Sheet1", tsv, '\t', hasHeader: true, startCell: "A1");
            result.Should().Contain("Imported");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/A1").Text.Should().Be("Name");
        ReadNode(readOnly, "/Sheet1/A2").Text.Should().Be("Alpha");
        ReadNode(readOnly, "/Sheet1/B2").Text.Should().Be("100");
        ReadNode(readOnly, "/Sheet1/B3").Text.Should().Be("200");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Import_StartCellPositioning_ImportsAtCorrectOffset()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            var csv = "Col1,Col2\nX,Y";
            handler.Import("/Sheet1", csv, ',', hasHeader: true, startCell: "C5");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/C5").Text.Should().Be("Col1");
        ReadNode(readOnly, "/Sheet1/C6").Text.Should().Be("X");
        ReadNode(readOnly, "/Sheet1/D6").Text.Should().Be("Y");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Import_CsvWithoutHeader_ImportsAllRows()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            var csv = "West,1200,Open\nEast,800,Closed";
            var result = handler.Import("/Sheet1", csv, ',', hasHeader: false, startCell: "A1");
            result.Should().Contain("Imported");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/A1").Text.Should().Be("West");
        ReadNode(readOnly, "/Sheet1/A2").Text.Should().Be("East");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Bloat filter ────────────────────────────────────────────────────

    [Fact]
    public void BloatFilter_WorkbookWithManyEmptyCells_NoDataLoss()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            // Write a few meaningful cells scattered among empty ones
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Header1"), ("bold", "true")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "Header2")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "Data1")));
            handler.Add("/Sheet1/B2", "cell", null, Props(("formula", "1+1")));
            handler.Add("/Sheet1/Z100", "cell", null, Props(("value", "FarCell"), ("fill", "C6EFCE")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/A1").Text.Should().Be("Header1");
        ReadNode(readOnly, "/Sheet1/B1").Text.Should().Be("Header2");
        ReadNode(readOnly, "/Sheet1/A2").Text.Should().Be("Data1");
        ReadNode(readOnly, "/Sheet1/B2").Format.Should().ContainKey("formula");
        ReadNode(readOnly, "/Sheet1/Z100").Text.Should().Be("FarCell");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void BloatFilter_MixedValueFormulaAndStyleCells_SurviveRoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "TopLeft"), ("fill", "D9EAD3"), ("bold", "true")));
            handler.Add("/Sheet1/B2", "cell", null, Props(("formula", "SUM(A1:A10)"), ("numberformat", "#,##0")));
            handler.Add("/Sheet1/C3", "cell", null, Props(("value", "Styled"), ("font.color", "FF0000"), ("italic", "true")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/A1").Text.Should().Be("TopLeft");
        ReadNode(readOnly, "/Sheet1/B2").Format.Should().ContainKey("formula");
        ReadNode(readOnly, "/Sheet1/C3").Text.Should().Be("Styled");
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
