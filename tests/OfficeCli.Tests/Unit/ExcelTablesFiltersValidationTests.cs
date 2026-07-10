// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class ExcelTablesFiltersValidationTests : ExcelTestBase
{
    [Fact]
    public void TablesAndDetectedTables_RoundTripThroughGetAndQuery()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales", "Status");
            AddRow(handler, 2, "West", "1200", "Open");
            AddRow(handler, 3, "East", "800", "Closed");
            handler.Add("/Sheet1", "table", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:C3",
                ["name"] = "SalesTable",
                ["displayName"] = "SalesTable",
                ["style"] = "medium2",
                ["headerRow"] = "true",
                ["totalRow"] = "true",
                ["showFirstColumn"] = "true",
                ["showLastColumn"] = "true",
                ["showRowStripes"] = "true",
                ["showColumnStripes"] = "true",
                ["columns"] = "Region,Sales,Status",
                ["totalsRowFunction"] = "none,sum,none"
            });

            handler.Add("/", "sheet", null, new(StringComparer.OrdinalIgnoreCase) { ["name"] = "Detected" });
            AddRow(handler, "Detected", 1, "Name", "Qty", "Price");
            AddRow(handler, "Detected", 2, "Pen", "2", "1.50");
            AddRow(handler, "Detected", 3, "Book", "1", "9.99");

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var table = ReadNode(readOnly, "/Sheet1/table[1]");
        table.Type.Should().Be("table");
        table.Text.Should().Be("SalesTable");
        table.Format.Should().Contain("name", "SalesTable");
        table.Format.Should().Contain("displayName", "SalesTable");
        table.Format.Should().Contain("ref", "A1:C4");
        table.Format.Should().Contain("style", "TableStyleMedium2");
        table.Format.Should().Contain("headerRow", true);
        table.Format.Should().Contain("totalRow", true);
        table.Format.Should().Contain("bandedRows", true);
        table.Format.Should().Contain("bandedCols", true);
        table.Format.Should().Contain("firstCol", true);
        table.Format.Should().Contain("lastCol", true);
        table.Format.Should().Contain("columns", "Region,Sales,Status");

        Query(readOnly, "listobject").Should().ContainSingle(n => n.Path == "/Sheet1/table[1]");
        Query(readOnly, "table").Should().Contain(n =>
            n.Type == "detectedtable" && n.Path == "/Detected/A1:C3");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void AutoFilterValidationAndSort_RoundTripAndAffectRows()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales", "Status");
            AddRow(handler, 2, "West", "1200", "Open");
            AddRow(handler, 3, "East", "800", "Closed");
            AddRow(handler, 4, "North", "1500", "Open");
            handler.Add("/Sheet1", "autofilter", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["range"] = "A1:C4",
                ["criteria0.equals"] = "West"
            });
            handler.Add("/Sheet1", "validation", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "C2:C4",
                ["type"] = "list",
                ["formula1"] = "Open,Closed",
                ["allowBlank"] = "false",
                ["showError"] = "true",
                ["showInput"] = "true",
                ["errorTitle"] = "Bad status",
                ["error"] = "Use list value",
                ["promptTitle"] = "Status",
                ["prompt"] = "Choose a status",
                ["errorStyle"] = "warning",
                ["inCellDropdown"] = "true"
            });
            handler.Set("/Sheet1", new(StringComparer.OrdinalIgnoreCase)
            {
                ["sort"] = "B desc",
                ["sortHeader"] = "true"
            });

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        sheet.Format.Should().Contain("autoFilter", "A1:C4");
        sheet.Format.Should().Contain("sort", "B:desc");
        ReadNode(readOnly, "/Sheet1/A2").Text.Should().Be("North");
        ReadNode(readOnly, "/Sheet1/A4").Text.Should().Be("East");

        var autoFilter = ReadNode(readOnly, "/Sheet1/autofilter");
        autoFilter.Type.Should().Be("autofilter");
        autoFilter.Format.Should().Contain("range", "A1:C4");

        var validation = Query(readOnly, "validation").Should().ContainSingle().Subject;
        validation.Path.Should().Be("/Sheet1/dataValidation[1]");
        validation.Format.Should().Contain("ref", "C2:C4");
        validation.Format.Should().Contain("type", "list");
        validation.Format.Should().Contain("formula1", "\"Open,Closed\"");
        validation.Format.Should().Contain("allowBlank", false);
        validation.Format.Should().Contain("showError", true);
        validation.Format.Should().Contain("showInput", true);
        validation.Format.Should().Contain("errorTitle", "Bad status");
        validation.Format.Should().Contain("error", "Use list value");
        validation.Format.Should().Contain("promptTitle", "Status");
        validation.Format.Should().Contain("prompt", "Choose a status");
        validation.Format.Should().Contain("errorStyle", "warning");
        validation.Format.Should().Contain("inCellDropdown", true);
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void NamedRanges_CanBeAddedUpdatedQueriedAndRemoved()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["value"] = "100" });
            handler.Add("/", "namedrange", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "Revenue",
                ["ref"] = "Sheet1!$A$1",
                ["scope"] = "workbook",
                ["comment"] = "Q4 total",
                ["volatile"] = "true"
            });
            handler.Set("/namedrange[Revenue]", new(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "Revenue2026",
                ["ref"] = "Sheet1!$A$1:$A$1",
                ["comment"] = "Updated total",
                ["volatile"] = "false"
            });
            handler.Save();
        }

        using (var readOnly = OpenReadOnly(path))
        {
            var namedRange = Query(readOnly, "namedrange").Should().ContainSingle().Subject;
            namedRange.Type.Should().Be("namedrange");
            namedRange.Format.Should().Contain("name", "Revenue2026");
            namedRange.Format.Should().Contain("ref", "Sheet1!$A$1:$A$1");
            namedRange.Format.Should().Contain("scope", "workbook");
            namedRange.Format.Should().Contain("comment", "Updated total");
            namedRange.Format.Should().NotContainKey("volatile");
            readOnly.Validate().Should().BeEmpty();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Remove("/namedrange[Revenue2026]", null);
            handler.Save();
        }

        using var afterRemove = OpenReadOnly(path);
        Query(afterRemove, "namedrange").Should().BeEmpty();
        afterRemove.Validate().Should().BeEmpty();
    }

    private static void AddRow(OfficeCli.Handlers.ExcelHandler handler, int row, params string[] values)
        => AddRow(handler, "Sheet1", row, values);

    private static void AddRow(OfficeCli.Handlers.ExcelHandler handler, string sheet, int row, params string[] values)
    {
        for (var i = 0; i < values.Length; i++)
        {
            var column = (char)('A' + i);
            handler.Add($"/{sheet}/{column}{row}", "cell", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["value"] = values[i]
            });
        }
    }
}
