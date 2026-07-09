// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelPivotTableUnitTests : ExcelTestBase
{
    [Fact]
    public void CreatePivotTable_AndReadback_ExposeCanonicalFieldValueFilterAndLayoutProperties()
    {
        var path = CreateWorkbookWithSourceData();

        using var document = SpreadsheetDocument.Open(path, true);
        var workbookPart = document.WorkbookPart!;
        var sheet = GetWorksheetPart(workbookPart, "Sheet1");

        PivotTableHelper.CreatePivotTable(
            workbookPart,
            sheet,
            sheet,
            "Sheet1",
            "A1:D5",
            "F1",
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "SalesPivot",
                ["rows"] = "Region",
                ["cols"] = "Category",
                ["filters"] = "Manager",
                ["values"] = "Sales:sum:percent_of_total:name=Sales Share",
                ["layout"] = "tabular",
                ["repeatLabels"] = "true",
                ["blankRows"] = "true",
                ["rowGrandTotals"] = "true",
                ["colGrandTotals"] = "false",
                ["subtotals"] = "off",
                ["style"] = "PivotStyleMedium9",
                ["showRowStripes"] = "true",
                ["showColHeaders"] = "false"
            }).Should().BeGreaterThan(0);

        var pivotPart = sheet.PivotTableParts.Should().ContainSingle().Subject;
        var node = new DocumentNode();
        PivotTableHelper.ReadPivotTableProperties(pivotPart.PivotTableDefinition!, node, pivotPart);

        node.Format.Should().Contain("name", "SalesPivot");
        node.Format.Should().Contain("source", "Sheet1!A1:D5");
        node.Format.Should().Contain("rows", "Region");
        node.Format.Should().Contain("cols", "Category");
        node.Format.Should().Contain("filters", "Manager");
        node.Format.Should().Contain("dataFieldCount", 1);
        node.Format.Should().Contain("dataField1.showAs", "percent_of_total");
        node.Format.Should().Contain("layout", "tabular");
        node.Format.Should().Contain("blankRows", "true");
        node.Format.Should().Contain("rowGrandTotals", "true");
        node.Format.Should().Contain("colGrandTotals", "false");
        node.Format.Should().Contain("subtotals", "off");
        node.Format.Should().Contain("style", "PivotStyleMedium9");
        node.Format.Should().Contain("showRowStripes", "true");
        node.Format.Should().Contain("showColHeaders", "false");

        var cacheDefinition = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!;
        cacheDefinition.RecordCount!.Value.Should().Be(4u);
        PivotTableHelper.NormalizePivotSource("'Sheet1'", "$A$1:$D$5").Should().Be("SHEET1!A1:D5");
        PivotTableHelper.ResolvePivotSourceSpec(workbookPart, "A1:D5", "Sheet1")
            .Should().Be(("Sheet1", "A1:D5"));
    }

    [Fact]
    public void SetPivotTableProperties_UpdatesReadbackForStyleFlagsAndGrandTotals()
    {
        var path = CreateWorkbookWithSourceData();

        using var document = SpreadsheetDocument.Open(path, true);
        var workbookPart = document.WorkbookPart!;
        var sheet = GetWorksheetPart(workbookPart, "Sheet1");

        PivotTableHelper.CreatePivotTable(
            workbookPart,
            sheet,
            sheet,
            "Sheet1",
            "A1:D5",
            "F1",
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "PivotToSet",
                ["rows"] = "Region",
                ["values"] = "Sales:sum"
            });

        var pivotPart = sheet.PivotTableParts.Should().ContainSingle().Subject;
        PivotTableHelper.SetPivotTableProperties(
            pivotPart,
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["repeatLabels"] = "true",
                ["blankRows"] = "true",
                ["rowGrandTotals"] = "false",
                ["colGrandTotals"] = "false",
                ["subtotals"] = "off",
                ["showColStripes"] = "true",
                ["showLastColumn"] = "false",
                ["style"] = "PivotStyleLight16"
            }).Should().BeEmpty();

        var node = new DocumentNode();
        PivotTableHelper.ReadPivotTableProperties(pivotPart.PivotTableDefinition!, node, pivotPart);

        node.Format.Should().Contain("repeatLabels", "true");
        node.Format.Should().Contain("blankRows", "true");
        node.Format.Should().Contain("rowGrandTotals", "false");
        node.Format.Should().Contain("colGrandTotals", "false");
        node.Format.Should().Contain("subtotals", "off");
        node.Format.Should().Contain("showColStripes", "true");
        node.Format.Should().Contain("showLastColumn", "false");
        node.Format.Should().Contain("style", "PivotStyleLight16");
    }

    [Fact]
    public void SharedCacheHelpers_ReusedCacheCanBeCountedFoundAndCloned()
    {
        var path = CreateWorkbookWithSourceData();

        using var document = SpreadsheetDocument.Open(path, true);
        var workbookPart = document.WorkbookPart!;
        var sheet = GetWorksheetPart(workbookPart, "Sheet1");
        var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["rows"] = "Region",
            ["values"] = "Sales:sum"
        };

        PivotTableHelper.CreatePivotTable(workbookPart, sheet, sheet, "Sheet1", "A1:D5", "F1", new(properties));
        PivotTableHelper.CreatePivotTable(workbookPart, sheet, sheet, "Sheet1", "A1:D5", "F20", new(properties));

        var pivotParts = sheet.PivotTableParts.ToList();
        pivotParts.Should().HaveCount(2);
        pivotParts[0].PivotTableCacheDefinitionPart.Should().BeSameAs(pivotParts[1].PivotTableCacheDefinitionPart);

        var sharedCache = pivotParts[0].PivotTableCacheDefinitionPart!;
        PivotTableHelper.FindMatchingCachePart(workbookPart, "Sheet1", "A1:D5").Should().BeSameAs(sharedCache);
        PivotTableHelper.CountCacheReferrers(workbookPart, sharedCache).Should().Be(2);

        var originalCacheId = PivotTableHelper.GetCacheIdForPart(workbookPart, sharedCache);
        originalCacheId.Should().NotBeNull();

        var clone = PivotTableHelper.CloneCachePartForCoW(workbookPart, sharedCache);
        clone.Should().NotBeSameAs(sharedCache);
        PivotTableHelper.GetCacheIdForPart(workbookPart, clone).Should().NotBe(originalCacheId);
        PivotTableHelper.CountCacheReferrers(workbookPart, clone).Should().Be(0);
        clone.GetPartsOfType<PivotTableCacheRecordsPart>().Should().ContainSingle();
    }

    private string CreateWorkbookWithSourceData()
    {
        var path = CreateWorkbook();
        using var handler = OpenEditable(path);

        AddRow(handler, 1, "Region", "Category", "Sales", "Manager");
        AddRow(handler, 2, "West", "Hardware", "1200", "Ada");
        AddRow(handler, 3, "East", "Software", "900", "Ben");
        AddRow(handler, 4, "West", "Software", "1500", "Ada");
        AddRow(handler, 5, "North", "Hardware", "600", "Cara");
        handler.Save();

        return path;
    }

    private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("Workbook is missing.");
        var sheets = workbook.Sheets
            ?? throw new InvalidOperationException("Workbook is missing sheets.");
        var sheet = sheets.Elements<Sheet>()
            .Single(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal));
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    }

    private static void AddRow(ExcelHandler handler, int row, params string[] values)
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
