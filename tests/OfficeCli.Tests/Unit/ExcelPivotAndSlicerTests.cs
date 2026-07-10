// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class ExcelPivotAndSlicerTests : ExcelTestBase
{
    [Fact]
    public void PivotTablesAndSlicers_RoundTripThroughGetAndQuery()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Category", "Sales", "Qty");
            AddRow(handler, 2, "West", "Hardware", "1200", "2");
            AddRow(handler, 3, "East", "Software", "800", "5");
            AddRow(handler, 4, "West", "Software", "1500", "3");
            AddRow(handler, 5, "North", "Hardware", "600", "1");

            handler.Add("/Sheet1", "pivottable", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["source"] = "Sheet1!A1:D5",
                ["position"] = "F1",
                ["name"] = "SalesPivot",
                ["style"] = "PivotStyleMedium9",
                ["rows"] = "Region",
                ["cols"] = "Category",
                ["values"] = "Sales:sum,Qty:count",
                ["layout"] = "tabular",
                ["repeatLabels"] = "true",
                ["blankRows"] = "true",
                ["rowGrandTotals"] = "true",
                ["colGrandTotals"] = "true",
                ["grandTotalCaption"] = "Total",
                ["subtotals"] = "off",
                ["showRowStripes"] = "true",
                ["showColStripes"] = "true"
            });
            handler.Add("/Sheet1", "slicer", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["pivotTable"] = "/Sheet1/pivottable[1]",
                ["field"] = "Region",
                ["name"] = "RegionSlicer",
                ["caption"] = "Filter Region",
                ["columnCount"] = "2",
                ["rowHeight"] = "250000",
                ["style"] = "SlicerStyleLight1"
            });

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var pivot = Query(readOnly, "pivottable").Should().ContainSingle().Subject;
        pivot.Path.Should().Be("/Sheet1/pivottable[1]");
        pivot.Format.Should().Contain("name", "SalesPivot");
        pivot.Format.Should().Contain("source", "Sheet1!A1:D5");
        pivot.Format.Should().Contain("style", "PivotStyleMedium9");
        pivot.Format.Should().Contain("rows", "Region");
        pivot.Format.Should().Contain("cols", "Category");
        pivot.Format.Should().Contain("dataFieldCount", 2);
        pivot.Format.Should().ContainKey("dataField1");
        pivot.Format.Should().ContainKey("dataField2");
        pivot.Format.Should().Contain("layout", "tabular");
        pivot.Format.Should().Contain("blankRows", "true");
        pivot.Format.Should().Contain("rowGrandTotals", "true");
        pivot.Format.Should().Contain("colGrandTotals", "true");
        pivot.Format.Should().Contain("grandTotalCaption", "Total");
        pivot.Format.Should().Contain("subtotals", "off");
        pivot.Format.Should().Contain("showRowStripes", "true");
        pivot.Format.Should().Contain("showColStripes", "true");

        var slicer = Query(readOnly, "slicer").Should().ContainSingle().Subject;
        slicer.Path.Should().Be("/Sheet1/slicer[1]");
        slicer.Format.Should().Contain("name", "RegionSlicer");
        slicer.Format.Should().Contain("caption", "Filter Region");
        slicer.Format.Should().Contain("field", "Region");
        slicer.Format.Should().Contain("columnCount", 2U);
        slicer.Format.Should().Contain("rowHeight", 250000U);
        slicer.Format.Should().Contain("style", "SlicerStyleLight1");
        slicer.Format.Should().ContainKey("pivotTable");
        slicer.Format.Should().ContainKey("pivotCacheId");
        slicer.Format.Should().Contain("itemCount", 3);
        readOnly.Validate().Should().BeEmpty();
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
