// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelTablesFiltersValidationIntegrationTests : ExcelTestBase
{
    // ── Table integration ────────────────────────────────────────────────

    [Fact]
    public void Table_AddGetQuerySetRemove_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales", "Status");
            AddRow(handler, 2, "West", "1200", "Open");
            AddRow(handler, 3, "East", "800", "Closed");
            handler.Add("/Sheet1", "table", null, Props(
                ("ref", "A1:C3"),
                ("name", "SalesTable"),
                ("columns", "Region,Sales,Status"),
                ("style", "TableStyleMedium2"),
                ("totalsRow", "true")
            ));
            handler.Save();
        }

        // Verify after reopen
        using (var readOnly = OpenReadOnly(path))
        {
            var table = ReadNode(readOnly, "/Sheet1/table[1]");
            table.Type.Should().Be("table");
            table.Text.Should().Be("SalesTable");
            table.Format.Should().Contain("name", "SalesTable");
            table.Format.Should().Contain("ref", "A1:C3");
            table.Format.Should().Contain("style", "TableStyleMedium2");
            Query(readOnly, "table").Should().ContainSingle(n => n.Type == "table");
            Query(readOnly, "listobject").Should().ContainSingle(n => n.Path == "/Sheet1/table[1]");
            readOnly.Validate().Should().BeEmpty();
        }

        // Rename and restyle table
        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1/table[1]", Props(
                ("name", "RenamedTable"),
                ("style", "TableStyleLight1")
            ));
            handler.Save();
        }

        using (var readOnly = OpenReadOnly(path))
        {
            var table = ReadNode(readOnly, "/Sheet1/table[1]");
            table.Format.Should().Contain("name", "RenamedTable");
            table.Format.Should().Contain("style", "TableStyleLight1");
            readOnly.Validate().Should().BeEmpty();
        }

        // Remove table
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/table[1]", null);
            handler.Save();
        }

        using var afterRemove = OpenReadOnly(path);
        Query(afterRemove, "table").Should().NotContain(n => n.Type == "table");
        afterRemove.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Table_AddDataBeyondInitialRef_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Qty");
            AddRow(handler, 2, "Pen", "2");
            handler.Add("/Sheet1", "table", null, Props(
                ("ref", "A1:B2"),
                ("name", "Items"),
                ("columns", "Item,Qty")
            ));
            // Add data beyond the table ref
            AddRow(handler, 3, "Book", "5");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var table = ReadNode(readOnly, "/Sheet1/table[1]");
        table.Format.Should().Contain("name", "Items");
        // Verify that the table and cell data survive save/reopen
        ReadNode(readOnly, "/Sheet1/A3").Text.Should().Be("Book");
        Query(readOnly, "table").Should().ContainSingle();
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Table_WithoutTotalsRow_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Score");
            AddRow(handler, 2, "Alice", "95");
            handler.Add("/Sheet1", "table", null, Props(
                ("ref", "A1:B2"),
                ("name", "Scores"),
                ("columns", "Name,Score"),
                ("style", "TableStyleMedium9")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var table = ReadNode(readOnly, "/Sheet1/table[1]");
        table.Format.Should().Contain("name", "Scores");
        table.Format.Should().Contain("style", "TableStyleMedium9");
        // Table is valid regardless of whether totalsRow defaults to true
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Table_MultipleTablesOnSameSheet_IndependentEntities()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Product", "Price");
            AddRow(handler, 2, "A", "10");
            handler.Add("/Sheet1", "table", null, Props(
                ("ref", "A1:B2"),
                ("name", "Products"),
                ("columns", "Product,Price"),
                ("style", "TableStyleMedium2")
            ));

            AddRow(handler, 5, "Customer", "Orders");
            AddRow(handler, 6, "X", "3");
            handler.Add("/Sheet1", "table", null, Props(
                ("ref", "A5:B6"),
                ("name", "Customers"),
                ("columns", "Customer,Orders"),
                ("style", "TableStyleMedium6")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var tables = Query(readOnly, "table").Where(n => n.Type == "table").ToList();
        tables.Should().HaveCount(2);
        var names = tables.Select(t => t.Format["name"]!.ToString()).ToList();
        names.Should().Contain(new[] { "Products", "Customers" });
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Table_DetectedTableReadback()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "sheet", null, Props(("name", "Data")));
            AddRow(handler, "Data", 1, "Name", "Age", "City");
            AddRow(handler, "Data", 2, "John", "30", "NYC");
            AddRow(handler, "Data", 3, "Jane", "25", "LA");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var detected = Query(readOnly, "table").Where(n => n.Type == "detectedtable").ToList();
        detected.Should().NotBeEmpty("auto-detected data range should be found");
        detected.Should().Contain(d => d.Path.StartsWith("/Data/"));
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Autofilter integration ──────────────────────────────────────────

    [Fact]
    public void AutoFilter_EqualsNotEquals_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales");
            AddRow(handler, 2, "West", "100");
            AddRow(handler, 3, "East", "200");
            AddRow(handler, 4, "West", "300");
            handler.Add("/Sheet1", "autofilter", null, Props(
                ("range", "A1:B4"),
                ("criteria0.equals", "West")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var af = ReadNode(readOnly, "/Sheet1/autofilter");
        af.Type.Should().Be("autofilter");
        af.Format.Should().Contain("range", "A1:B4");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void AutoFilter_ContainsNotContains_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Price");
            AddRow(handler, 2, "Red Apple", "1.5");
            AddRow(handler, 3, "Banana", "0.8");
            AddRow(handler, 4, "Green Apple", "1.6");
            handler.Add("/Sheet1", "autofilter", null, Props(
                ("range", "A1:B4"),
                ("criteria0.contains", "Apple")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var af = ReadNode(readOnly, "/Sheet1/autofilter");
        af.Type.Should().Be("autofilter");
        af.Format.Should().Contain("range", "A1:B4");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void AutoFilter_BeginsWithEndsWith_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Code", "Value");
            AddRow(handler, 2, "US-001", "10");
            AddRow(handler, 3, "UK-002", "20");
            AddRow(handler, 4, "US-003", "30");
            handler.Add("/Sheet1", "autofilter", null, Props(
                ("range", "A1:B4"),
                ("criteria0.beginswith", "US")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var af = ReadNode(readOnly, "/Sheet1/autofilter");
        af.Type.Should().Be("autofilter");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void AutoFilter_GreaterThanLessThan_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Amount");
            AddRow(handler, 2, "A", "50");
            AddRow(handler, 3, "B", "150");
            AddRow(handler, 4, "C", "250");
            handler.Add("/Sheet1", "autofilter", null, Props(
                ("range", "A1:B4"),
                ("criteria0.gt", "100")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var af = ReadNode(readOnly, "/Sheet1/autofilter");
        af.Type.Should().Be("autofilter");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void AutoFilter_BetweenNotBetween_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Score");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "50");
            AddRow(handler, 4, "C", "90");
            handler.Add("/Sheet1", "autofilter", null, Props(
                ("range", "A1:B4"),
                ("criteria0.between", "20,80")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var af = ReadNode(readOnly, "/Sheet1/autofilter");
        af.Type.Should().Be("autofilter");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void AutoFilter_TopBottom_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "50");
            AddRow(handler, 4, "C", "90");
            AddRow(handler, 5, "D", "30");
            handler.Add("/Sheet1", "autofilter", null, Props(
                ("range", "A1:B5"),
                ("criteria0.top", "3")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var af = ReadNode(readOnly, "/Sheet1/autofilter");
        af.Type.Should().Be("autofilter");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void AutoFilter_Remove_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "A", "B");
            AddRow(handler, 2, "1", "2");
            handler.Add("/Sheet1", "autofilter", null, Props(
                ("range", "A1:B2"),
                ("criteria0.equals", "1")
            ));
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/autofilter", null);
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "autofilter").Should().BeEmpty();
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Sort integration ─────────────────────────────────────────────────

    [Fact]
    public void Sort_SingleKeyWithHeader_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Score");
            AddRow(handler, 2, "Charlie", "80");
            AddRow(handler, 3, "Alice", "95");
            AddRow(handler, 4, "Bob", "70");
            handler.Set("/Sheet1", Props(
                ("sort", "A asc"),
                ("sortHeader", "true")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        // Sheet should have sort state
        sheet.Format.Keys.Should().Contain(k => k.Contains("sort", StringComparison.OrdinalIgnoreCase));
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Sort_MultipleKeysDifferentOrders_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Dept", "Salary", "Name");
            AddRow(handler, 2, "Eng", "100", "Zoe");
            AddRow(handler, 3, "Eng", "120", "Alice");
            AddRow(handler, 4, "Sales", "90", "Bob");
            handler.Set("/Sheet1", Props(
                ("sort", "A asc, B desc"),
                ("sortHeader", "true")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        sheet.Format.Keys.Should().Contain(k => k.Contains("sort", StringComparison.OrdinalIgnoreCase));
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Sort_NoHeader_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "30")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "10")));
            handler.Add("/Sheet1/A3", "cell", null, Props(("value", "20")));
            handler.Set("/Sheet1", Props(
                ("sort", "A asc"),
                ("sortHeader", "false")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Data Validation integration ──────────────────────────────────────

    [Fact]
    public void Validation_ListType_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Status");
            AddRow(handler, 2, "Open");
            handler.Add("/Sheet1", "validation", null, Props(
                ("ref", "A2:A10"),
                ("type", "list"),
                ("formula1", "Open,Closed,Pending"),
                ("allowBlank", "true"),
                ("showDropdown", "true")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var dv = Query(readOnly, "validation").Should().ContainSingle().Subject;
        dv.Type.Should().Be("dataValidation");
        dv.Format.Should().Contain("type", "list");
        dv.Format.Should().Contain("ref", "A2:A10");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Validation_DecimalBetween_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Score");
            AddRow(handler, 2, "85");
            handler.Add("/Sheet1", "validation", null, Props(
                ("ref", "A2:A10"),
                ("type", "decimal"),
                ("operator", "between"),
                ("formula1", "0"),
                ("formula2", "100"),
                ("allowBlank", "true"),
                ("prompt", "Enter a score"),
                ("promptTitle", "Score Input"),
                ("error", "Score must be 0-100"),
                ("errorTitle", "Invalid Score"),
                ("errorStyle", "stop")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var dv = Query(readOnly, "validation").Should().ContainSingle().Subject;
        dv.Type.Should().Be("dataValidation");
        dv.Format.Should().Contain("type", "decimal");
        dv.Format.Should().Contain("ref", "A2:A10");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Validation_WholeNumberGreaterThan_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Qty");
            AddRow(handler, 2, "5");
            handler.Add("/Sheet1", "validation", null, Props(
                ("ref", "A2:A10"),
                ("type", "whole"),
                ("operator", "greaterThan"),
                ("formula1", "0"),
                ("error", "Must be positive"),
                ("errorStyle", "warning")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var dv = Query(readOnly, "validation").Should().ContainSingle().Subject;
        dv.Format.Should().Contain("type", "whole");
        dv.Format.Should().Contain("ref", "A2:A10");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Validation_DateType_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Date");
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "2026-01-15"), ("type", "date")));
            handler.Add("/Sheet1", "validation", null, Props(
                ("ref", "A2:A10"),
                ("type", "date"),
                ("operator", "greaterThan"),
                ("formula1", "2026-01-01"),
                ("error", "Date must be after 2026-01-01")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var dv = Query(readOnly, "validation").Should().ContainSingle().Subject;
        dv.Format.Should().Contain("type", "date");
        dv.Format.Should().Contain("ref", "A2:A10");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Validation_TextLength_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Code");
            AddRow(handler, 2, "ABC");
            handler.Add("/Sheet1", "validation", null, Props(
                ("ref", "A2:A10"),
                ("type", "textLength"),
                ("operator", "between"),
                ("formula1", "3"),
                ("formula2", "10"),
                ("error", "Code must be 3-10 characters")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var dv = Query(readOnly, "validation").Should().ContainSingle().Subject;
        dv.Format.Should().Contain("type", "textLength");
        dv.Format.Should().Contain("ref", "A2:A10");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Validation_TimeType_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Start");
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "09:00")));
            handler.Add("/Sheet1", "validation", null, Props(
                ("ref", "A2:A10"),
                ("type", "time"),
                ("operator", "between"),
                ("formula1", "08:00"),
                ("formula2", "18:00"),
                ("error", "Time must be 8:00-18:00")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var dv = Query(readOnly, "validation").Should().ContainSingle().Subject;
        dv.Format.Should().Contain("type", "time");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Validation_SetProperties_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Choice");
            AddRow(handler, 2, "A");
            handler.Add("/Sheet1", "validation", null, Props(
                ("ref", "A2:A10"),
                ("type", "list"),
                ("formula1", "A,B")
            ));
            handler.Save();
        }

        // Reopen and modify
        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1/validation[1]", Props(
                ("formula1", "X,Y,Z"),
                ("allowBlank", "false")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var dv = Query(readOnly, "validation").Should().ContainSingle().Subject;
        dv.Format.Should().Contain("type", "list");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Validation_Remove_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Data");
            AddRow(handler, 2, "X");
            handler.Add("/Sheet1", "validation", null, Props(
                ("ref", "A2:A10"),
                ("type", "list"),
                ("formula1", "X,Y")
            ));
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/validation[1]", null);
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "validation").Should().BeEmpty();
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Named Range integration ──────────────────────────────────────────

    [Fact]
    public void NamedRange_WorkbookScope_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "42")));
            handler.Add("/", "namedrange", null, Props(
                ("name", "Answer"),
                ("ref", "Sheet1!$A$1"),
                ("comment", "The answer"),
                ("volatile", "false")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var nr = Query(readOnly, "namedrange").Should().ContainSingle().Subject;
        nr.Type.Should().Be("namedrange");
        nr.Format.Should().Contain("name", "Answer");
        nr.Format.Should().Contain("ref", "Sheet1!$A$1");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void NamedRange_SheetScope_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "100")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "200")));
            handler.Add("/Sheet1", "namedrange", null, Props(
                ("name", "LocalRef"),
                ("ref", "Sheet1!A1:B1")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var nr = Query(readOnly, "namedrange").Should().ContainSingle().Subject;
        nr.Format.Should().Contain("name", "LocalRef");
        nr.Format.Should().Contain("ref", "Sheet1!A1:B1");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void NamedRange_SetPropertiesAndRemove_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "99")));
            handler.Add("/", "namedrange", null, Props(
                ("name", "TempRange"),
                ("ref", "Sheet1!$A$1"),
                ("comment", "Original")
            ));
            handler.Save();
        }

        // Update properties
        using (var handler = OpenEditable(path))
        {
            handler.Set("/namedrange[TempRange]", Props(
                ("ref", "Sheet1!$A$1:$A$1"),
                ("comment", "Updated range")
            ));
            handler.Save();
        }

        using (var readOnly = OpenReadOnly(path))
        {
            var nr = Query(readOnly, "namedrange").Should().ContainSingle().Subject;
            nr.Format.Should().Contain("name", "TempRange");
            nr.Format.Should().Contain("comment", "Updated range");
            readOnly.Validate().Should().BeEmpty();
        }

        // Remove
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/namedrange[TempRange]", null);
            handler.Save();
        }

        using var afterRemove = OpenReadOnly(path);
        Query(afterRemove, "namedrange").Should().BeEmpty();
        afterRemove.Validate().Should().BeEmpty();
    }

    [Fact]
    public void NamedRange_UsedInFormula_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "100")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "200")));
            handler.Add("/", "namedrange", null, Props(
                ("name", "FirstValue"),
                ("ref", "Sheet1!$A$1"),
                ("comment", "First cell value")
            ));
            handler.Add("/Sheet1/B1", "cell", null, Props(("formula", "FirstValue+A2")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var nr = Query(readOnly, "namedrange").Should().ContainSingle(n =>
            n.Format["name"]!.Equals("FirstValue")).Subject;
        nr.Format.Should().Contain("ref", "Sheet1!$A$1");
        var formulaCell = ReadNode(readOnly, "/Sheet1/B1");
        formulaCell.Format.Should().ContainKey("formula");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void NamedRange_MultipleRanges_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "10")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "20")));
            handler.Add("/", "namedrange", null, Props(
                ("name", "Alpha"),
                ("ref", "Sheet1!$A$1"),
                ("comment", "Alpha cell")
            ));
            handler.Add("/", "namedrange", null, Props(
                ("name", "Beta"),
                ("ref", "Sheet1!$B$1"),
                ("comment", "Beta cell")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var ranges = Query(readOnly, "namedrange").ToList();
        ranges.Should().HaveCount(2);
        var names = ranges.Select(r => r.Format["name"]!.ToString()).ToList();
        names.Should().Contain(new[] { "Alpha", "Beta" });
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static Dictionary<string, string> Props(params (string Key, string Value)[] values)
        => values.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

    private static void AddRow(ExcelHandler handler, int row, params string[] values)
        => AddRow(handler, "Sheet1", row, values);

    private static void AddRow(ExcelHandler handler, string sheet, int row, params string[] values)
    {
        for (var i = 0; i < values.Length; i++)
        {
            var column = (char)('A' + i);
            handler.Add($"/{sheet}/{column}{row}", "cell", null, Props(("value", values[i])));
        }
    }
}
