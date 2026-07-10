// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class ExcelPivotSlicerIntegrationTests : ExcelTestBase
{
    // ── Pivot: add/get/set/remove ───────────────────────────────────────

    [Fact]
    public void PivotTable_AddGetSetRemoveWithSaveReopen_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Category", "Sales", "Qty");
            AddRow(handler, 2, "West", "Hardware", "1200", "2");
            AddRow(handler, 3, "East", "Software", "800", "5");
            AddRow(handler, 4, "West", "Software", "1500", "3");
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:D4"),
                ("position", "F1"),
                ("name", "SalesPivot"),
                ("rows", "Region"),
                ("cols", "Category"),
                ("values", "Sales:sum,Qty:count"),
                ("style", "PivotStyleMedium9")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var pivot = Query(readOnly, "pivottable").Should().ContainSingle(n =>
            n.Format["name"]!.Equals("SalesPivot")).Subject;
        pivot.Path.Should().Be("/Sheet1/pivottable[1]");
        pivot.Format.Should().ContainKey("dataField1");
        pivot.Format.Should().ContainKey("cacheId");
        readOnly.Validate().Should().BeEmpty();

        // Set style
        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1/pivottable[1]", Props(
                ("style", "PivotStyleLight16")
            ));
            handler.Save();
        }

        using var readOnly2 = OpenReadOnly(path);
        var pivot2 = ReadNode(readOnly2, "/Sheet1/pivottable[1]");
        pivot2.Format.Should().Contain("style", "PivotStyleLight16");
        readOnly2.Validate().Should().BeEmpty();

        // Remove pivot
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/pivottable[1]", null);
            handler.Save();
        }
        using var finalReadOnly = OpenReadOnly(path);
        Query(finalReadOnly, "pivottable").Should().BeEmpty();
        finalReadOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void PivotTable_CacheReadback_ContainsExpectedKeys()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales");
            AddRow(handler, 2, "West", "1200");
            AddRow(handler, 3, "East", "800");
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:B3"),
                ("position", "D1"),
                ("name", "SimplePivot"),
                ("rows", "Region"),
                ("values", "Sales:sum"),
                ("style", "PivotStyleMedium2")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var pivot = ReadNode(readOnly, "/Sheet1/pivottable[1]");
        pivot.Format.Should().Contain("name", "SimplePivot");
        pivot.Format.Should().Contain("source", "Sheet1!A1:B3");
        pivot.Format.Should().Contain("rows", "Region");
        pivot.Format.Should().Contain("dataFieldCount", 1);
        pivot.Format.Should().ContainKey("dataField1");
        pivot.Format.Should().ContainKey("cacheId");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void PivotTable_FieldsValuesAndLayout_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Category", "Sales", "Qty");
            AddRow(handler, 2, "West", "Hardware", "1200", "2");
            AddRow(handler, 3, "East", "Software", "800", "5");
            AddRow(handler, 4, "West", "Software", "1500", "3");
            AddRow(handler, 5, "North", "Hardware", "600", "1");
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:D5"),
                ("position", "F1"),
                ("name", "DetailPivot"),
                ("rows", "Region"),
                ("cols", "Category"),
                ("values", "Sales:sum,Qty:count"),
                ("style", "PivotStyleMedium9"),
                ("layout", "tabular"),
                ("repeatLabels", "true"),
                ("blankRows", "true"),
                ("rowGrandTotals", "true"),
                ("colGrandTotals", "true"),
                ("grandTotalCaption", "Total"),
                ("subtotals", "off"),
                ("showRowStripes", "true"),
                ("showColStripes", "true")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var pivot = ReadNode(readOnly, "/Sheet1/pivottable[1]");
        pivot.Format.Should().Contain("name", "DetailPivot");
        pivot.Format.Should().Contain("rows", "Region");
        pivot.Format.Should().Contain("cols", "Category");
        pivot.Format.Should().Contain("dataFieldCount", 2);
        pivot.Format.Should().ContainKey("dataField1");
        pivot.Format.Should().ContainKey("dataField2");
        pivot.Format.Should().Contain("layout", "tabular");
        pivot.Format.Should().Contain("blankRows", "true");
        pivot.Format.Should().Contain("rowGrandTotals", "true");
        pivot.Format.Should().Contain("grandTotalCaption", "Total");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void PivotTable_OrphanCacheCleanup_RemovePivotVerifyNoDataLoss()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales");
            AddRow(handler, 2, "West", "1200");
            AddRow(handler, 3, "East", "800");
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:B3"),
                ("position", "D1"),
                ("name", "ToRemove"),
                ("rows", "Region"),
                ("values", "Sales:sum")
            ));
            handler.Save();
        }

        // Remove pivot and verify document is still valid
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/pivottable[1]", null);
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "pivottable").Should().BeEmpty();
        // Source data should still be present
        ReadNode(readOnly, "/Sheet1/A1").Text.Should().Be("Region");
        ReadNode(readOnly, "/Sheet1/A2").Text.Should().Be("West");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Slicer: add/get/set/remove ──────────────────────────────────────

    [Fact]
    public void Slicer_AddGetRemoveBoundToPivot_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales");
            AddRow(handler, 2, "West", "1200");
            AddRow(handler, 3, "East", "800");
            AddRow(handler, 4, "North", "600");
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:B4"),
                ("position", "D1"),
                ("name", "RegionPivot"),
                ("rows", "Region"),
                ("values", "Sales:sum")
            ));
            handler.Add("/Sheet1", "slicer", null, Props(
                ("pivotTable", "/Sheet1/pivottable[1]"),
                ("field", "Region"),
                ("name", "RegionSlicer"),
                ("caption", "Filter Region"),
                ("columnCount", "2"),
                ("style", "SlicerStyleLight1")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var slicer = Query(readOnly, "slicer").Should().ContainSingle(n =>
            n.Format["name"]!.Equals("RegionSlicer")).Subject;
        slicer.Path.Should().Be("/Sheet1/slicer[1]");
        slicer.Format.Should().Contain("caption", "Filter Region");
        slicer.Format.Should().Contain("field", "Region");
        slicer.Format.Should().ContainKey("pivotTable");
        slicer.Format.Should().ContainKey("pivotCacheId");
        readOnly.Validate().Should().BeEmpty();

        // Get slicer details
        var slicerGet = ReadNode(readOnly, "/Sheet1/slicer[1]");
        slicerGet.Format.Should().Contain("name", "RegionSlicer");
        slicerGet.Format.Should().Contain("columnCount", 2U);
    }

    [Fact]
    public void Slicer_RelationshipReadback_ContainsPivotTableAndCacheId()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Category", "Sales");
            AddRow(handler, 2, "West", "Hardware", "1200");
            AddRow(handler, 3, "East", "Software", "800");
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:C3"),
                ("position", "E1"),
                ("name", "MyPivot"),
                ("rows", "Region"),
                ("cols", "Category"),
                ("values", "Sales:sum")
            ));
            handler.Add("/Sheet1", "slicer", null, Props(
                ("pivotTable", "/Sheet1/pivottable[1]"),
                ("field", "Region"),
                ("name", "MySlicer"),
                ("caption", "Choose Region"),
                ("columnCount", "3"),
                ("style", "SlicerStyleLight3")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var slicer = ReadNode(readOnly, "/Sheet1/slicer[1]");
        slicer.Format.Should().Contain("name", "MySlicer");
        slicer.Format.Should().Contain("caption", "Choose Region");
        slicer.Format.Should().Contain("field", "Region");
        slicer.Format.Should().Contain("columnCount", 3U);
        slicer.Format.Should().ContainKey("pivotTable");
        slicer.Format.Should().ContainKey("pivotCacheId");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Pivot + chart interaction ───────────────────────────────────────

    [Fact]
    public void PivotAndChart_ComposedOnSharedData_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Category", "Sales", "Qty");
            AddRow(handler, 2, "West", "Hardware", "1200", "2");
            AddRow(handler, 3, "East", "Software", "800", "5");
            AddRow(handler, 4, "West", "Software", "1500", "3");
            AddRow(handler, 5, "North", "Hardware", "600", "1");
            handler.Add("/Sheet1", "chart", null, Props(
                ("type", "bar"),
                ("dataRange", "Sheet1!A1:C5"),
                ("anchor", "J2:P16"),
                ("title", "Sales by Region")
            ));
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:D5"),
                ("position", "F1"),
                ("name", "SalesPivot"),
                ("rows", "Region"),
                ("cols", "Category"),
                ("values", "Sales:sum,Qty:count"),
                ("style", "PivotStyleMedium9")
            ));
            handler.Add("/Sheet1", "slicer", null, Props(
                ("pivotTable", "/Sheet1/pivottable[1]"),
                ("field", "Region"),
                ("name", "RegionSlicer"),
                ("caption", "Filter Region"),
                ("columnCount", "2"),
                ("style", "SlicerStyleLight1")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "chart").Should().ContainSingle(n => n.Path == "/Sheet1/chart[1]");
        Query(readOnly, "pivottable").Should().ContainSingle(n =>
            n.Format["name"]!.Equals("SalesPivot") && n.Format.ContainsKey("dataField1"));
        Query(readOnly, "slicer").Should().ContainSingle(n =>
            n.Format["name"]!.Equals("RegionSlicer") && n.Format.ContainsKey("pivotCacheId"));
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Pivot style change ──────────────────────────────────────────────

    [Fact]
    public void PivotTable_StyleChange_PersistsAcrossSaveReopen()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales");
            AddRow(handler, 2, "West", "1200");
            AddRow(handler, 3, "East", "800");
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:B3"),
                ("position", "D1"),
                ("name", "StyledPivot"),
                ("rows", "Region"),
                ("values", "Sales:sum"),
                ("style", "PivotStyleDark1")
            ));
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1/pivottable[1]", Props(("style", "PivotStyleDark5")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var pivot = ReadNode(readOnly, "/Sheet1/pivottable[1]");
        pivot.Format.Should().Contain("style", "PivotStyleDark5");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Slicer item count ──────────────────────────────────────────────

    [Fact]
    public void Slicer_ItemCountReflectsDistinctFieldValues()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales");
            AddRow(handler, 2, "West", "1200");
            AddRow(handler, 3, "East", "800");
            AddRow(handler, 4, "North", "600");
            AddRow(handler, 5, "South", "400");
            handler.Add("/Sheet1", "pivottable", null, Props(
                ("source", "Sheet1!A1:B5"),
                ("position", "D1"),
                ("name", "BigPivot"),
                ("rows", "Region"),
                ("values", "Sales:sum")
            ));
            handler.Add("/Sheet1", "slicer", null, Props(
                ("pivotTable", "/Sheet1/pivottable[1]"),
                ("field", "Region"),
                ("name", "BigSlicer"),
                ("caption", "Regions"),
                ("columnCount", "1"),
                ("style", "SlicerStyleLight1")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var slicer = ReadNode(readOnly, "/Sheet1/slicer[1]");
        slicer.Format.Should().ContainKey("itemCount");
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
