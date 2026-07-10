// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class ExcelCellsAndFormulasTests : ExcelTestBase
{
    [Fact]
    public void CreateWorkbook_AddCellsAndStyles_ReadsBackAfterSave()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["value"] = "Hello",
                ["bold"] = "true",
                ["italic"] = "true",
                ["font.color"] = "FF0000",
                ["fill"] = "FFFF00",
                ["numberformat"] = "@",
                ["alignment.horizontal"] = "center",
                ["alignment.vertical"] = "center",
                ["alignment.wrapText"] = "true"
            });
            handler.Add("/Sheet1/B1", "cell", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["value"] = "42",
                ["type"] = "number",
                ["numberformat"] = "#,##0.00"
            });
            handler.Add("/Sheet1/C1", "cell", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["value"] = "2026-07-08",
                ["type"] = "date",
                ["numberformat"] = "yyyy-mm-dd"
            });
            handler.Add("/Sheet1/D1", "cell", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["value"] = "true",
                ["type"] = "boolean"
            });

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var textCell = ReadNode(readOnly, "/Sheet1/A1");
        textCell.Text.Should().Be("Hello");
        textCell.Format.Should().Contain("font.bold", true);
        textCell.Format.Should().Contain("font.italic", true);
        textCell.Format.Should().Contain("font.color", "#FF0000");
        textCell.Format.Should().Contain("fill", "#FFFF00");
        textCell.Format.Should().Contain("numberformat", "@");
        textCell.Format.Should().Contain("alignment.horizontal", "center");
        textCell.Format.Should().Contain("alignment.vertical", "center");
        textCell.Format.Should().Contain("alignment.wrapText", true);

        ReadNode(readOnly, "/Sheet1/B1").Text.Should().Be("42");
        ReadNode(readOnly, "/Sheet1/C1").Format.Should().Contain("numberformat", "yyyy-mm-dd");
        var booleanCell = ReadNode(readOnly, "/Sheet1/D1");
        booleanCell.Text.Should().Be("1");
        booleanCell.Format.Should().Contain("type", "Boolean");
        Query(readOnly, "cell").Select(n => n.Path).Should().Contain(new[] { "/Sheet1/A1", "/Sheet1/B1", "/Sheet1/C1", "/Sheet1/D1" });
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void SheetRowColumnRangeAndMerge_SupportAddSetMoveAndRemove()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "sheet", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "Data",
                ["tabColor"] = "4472C4",
                ["hidden"] = "true"
            });
            handler.Set("/Data", new(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "Summary",
                ["freeze"] = "B2",
                ["direction"] = "rtl",
                ["header"] = "&CQuarterly",
                ["footer"] = "&P"
            });
            handler.Set("/Summary", new(StringComparer.OrdinalIgnoreCase)
            {
                ["printArea"] = "A1:C20"
            });
            handler.Add("/Summary", "row", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["cols"] = "3",
                ["height"] = "24",
                ["hidden"] = "true"
            });
            handler.Add("/Summary", "col", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "B",
                ["width"] = "18",
                ["hidden"] = "true"
            });
            handler.Add("/Summary/A3", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["value"] = "A" });
            handler.Add("/Summary/B3", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["value"] = "B" });
            handler.Add("/Summary/C3", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["value"] = "C" });
            handler.Move("/Summary/col[A]", null, InsertPosition.AfterElement("/Summary/col[C]"));
            handler.Remove("/Summary/C3", null);
            handler.Add("/Summary/A1", "cell", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["value"] = "Merged",
                ["merge"] = "A1:B2"
            });

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Summary");
        sheet.Path.Should().Be("/Summary");
        sheet.Format.Should().Contain("tabColor", "#4472C4");
        sheet.Format.Should().Contain("hidden", true);
        sheet.Format.Should().Contain("visibility", "hidden");
        sheet.Format.Should().Contain("freeze", "B2");
        sheet.Format.Should().Contain("direction", "rtl");
        sheet.Format.Should().Contain("printArea", "A1:D20");
        sheet.Format.Should().Contain("header", "&CQuarterly");
        sheet.Format.Should().Contain("footer", "&P");

        var row = ReadNode(readOnly, "/Summary/row[1]");
        row.Format.Should().Contain("height", "24pt");
        row.Format.Should().Contain("hidden", true);

        var column = ReadNode(readOnly, "/Summary/col[A]");
        column.Format.Should().Contain("width", 18D);
        column.Format.Should().Contain("hidden", true);

        var merged = ReadNode(readOnly, "/Summary/A1");
        merged.Text.Should().Be("Merged");
        merged.Format.Should().Contain("merge", "A1:B2");
        ReadNode(readOnly, "/Summary/A1:B2").Type.Should().Be("range");
        Query(readOnly, "sheet").Select(n => n.Path).Should().Contain("/Summary");
        Query(readOnly, "row").Select(n => n.Path).Should().Contain("/Summary/row[1]");
        Query(readOnly, "column").Select(n => n.Path).Should().Contain("/Summary/col[A]");
        Query(readOnly, "cell").Select(n => n.Path).Should().NotContain("/Summary/D1");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void RichTextRunsCommentsAndHyperlinks_RoundTripThroughQueryAndGet()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["value"] = "Intro" });
            handler.Add("/Sheet1/A1", "run", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["text"] = " Bold",
                ["bold"] = "true",
                ["font.color"] = "00AA00"
            });
            handler.Add("/Sheet1/B2", "comment", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["text"] = "Review formula",
                ["author"] = "OfficeCLI",
                ["font.bold"] = "true",
                ["font.color"] = "FF0000"
            });
            handler.Add("/Sheet1/C3", "cell", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["value"] = "Docs",
                ["link"] = "https://officecli.ai",
                ["tooltip"] = "Open docs",
                ["display"] = "OfficeCLI Docs"
            });

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var richCell = ReadNode(readOnly, "/Sheet1/A1");
        richCell.Text.Should().Be(" Bold");
        richCell.Format.Should().Contain("richtext", true);
        richCell.ChildCount.Should().Be(1);
        var run = ReadNode(readOnly, "/Sheet1/A1/run[1]");
        run.Text.Should().Be(" Bold");
        run.Format.Should().Contain("bold", true);
        run.Format.Should().Contain("color", "#00AA00");

        var comment = Query(readOnly, "comment").Should().ContainSingle().Subject;
        comment.Path.Should().Be("/Sheet1/comment[1]");
        comment.Text.Should().Be("Review formula");
        comment.Format.Should().Contain("ref", "B2");
        comment.Format.Should().Contain("author", "OfficeCLI");
        comment.Format.Should().Contain("font.bold", true);
        comment.Format.Should().Contain("font.color", "#FF0000");

        var linkedCell = ReadNode(readOnly, "/Sheet1/C3");
        linkedCell.Format.Should().Contain("link", "https://officecli.ai");
        linkedCell.Format.Should().Contain("tooltip", "Open docs");
        linkedCell.Format.Should().Contain("display", "OfficeCLI Docs");
        var hyperlink = Query(readOnly, "hyperlink").Should().ContainSingle().Subject;
        hyperlink.Path.Should().Be("/Sheet1/C3");
        hyperlink.Format.Should().Contain("url", "https://officecli.ai/");
        hyperlink.Format.Should().Contain("tooltip", "Open docs");
        hyperlink.Format.Should().Contain("display", "OfficeCLI Docs");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void FormulaFamiliesAndWorkbookProperties_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/", new(StringComparer.OrdinalIgnoreCase)
            {
                ["author"] = "OfficeCLI Tests",
                ["title"] = "Excel Coverage",
                ["calc.mode"] = "manual",
                ["calc.iterate"] = "true",
                ["calc.iterateCount"] = "25",
                ["calc.iterateDelta"] = "0.001",
                ["calc.fullPrecision"] = "true"
            });
            handler.Add("/Sheet1/A1", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["value"] = "1" });
            handler.Add("/Sheet1/A2", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["value"] = "2" });
            handler.Add("/Sheet1/A3", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["value"] = "3" });
            handler.Add("/Sheet1/B1", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["formula"] = "SUM(A1:A3)" });
            handler.Add("/Sheet1/B2", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["formula"] = "TEXT(DATE(2026,7,8),\"yyyy-mm-dd\")" });
            handler.Add("/Sheet1/B3", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["formula"] = "VLOOKUP(2,A1:A3,1,FALSE)" });
            handler.Add("/Sheet1/B4", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["formula"] = "SEQUENCE(2,2)" });
            handler.Add("/Sheet1/B5", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["formula"] = "PMT(0.05/12,12,1000)" });
            handler.Add("/Sheet1/B6", "cell", null, new(StringComparer.OrdinalIgnoreCase) { ["formula"] = "STDEV.S(A1:A3)" });

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var root = ReadNode(readOnly, "/");
        root.Format.Should().Contain("author", "OfficeCLI Tests");
        root.Format.Should().Contain("title", "Excel Coverage");
        root.Format.Should().Contain("calc.mode", "manual");
        root.Format.Should().Contain("calc.iterate", true);
        root.Format.Should().Contain("calc.iterateCount", 25U);
        root.Format.Should().Contain("calc.iterateDelta", 0.001D);
        root.Format.Should().Contain("calc.fullPrecision", true);

        ReadNode(readOnly, "/Sheet1/B1").Format.Should().Contain("formula", "SUM(A1:A3)");
        ReadNode(readOnly, "/Sheet1/B2").Format.Should().Contain("formula", "TEXT(DATE(2026,7,8),\"yyyy-mm-dd\")");
        ReadNode(readOnly, "/Sheet1/B3").Format.Should().Contain("formula", "VLOOKUP(2,A1:A3,1,FALSE)");
        var dynamicArray = ReadNode(readOnly, "/Sheet1/B4");
        dynamicArray.Format.Should().Contain("formula", "SEQUENCE(2,2)");
        dynamicArray.Format.Should().Contain("arrayformula", true);
        ReadNode(readOnly, "/Sheet1/B5").Format.Should().Contain("formula", "PMT(0.05/12,12,1000)");
        ReadNode(readOnly, "/Sheet1/B6").Format.Should().Contain("formula", "STDEV.S(A1:A3)");
        readOnly.Validate().Should().BeEmpty();
    }
}
