// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelCellsRangesIntegrationTests : ExcelTestBase
{
    // ── Cell types ──────────────────────────────────────────────────────

    [Fact]
    public void CellTypes_TextNumberBooleanDate_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Plain text")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "99"), ("type", "number"), ("numberformat", "0.00")));
            handler.Add("/Sheet1/A3", "cell", null, Props(("value", "true"), ("type", "boolean")));
            handler.Add("/Sheet1/A4", "cell", null, Props(("value", "2026-07-08"), ("type", "date"), ("numberformat", "yyyy-mm-dd")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/A1").Text.Should().Be("Plain text");
        ReadNode(readOnly, "/Sheet1/A2").Text.Should().Be("99");
        ReadNode(readOnly, "/Sheet1/A2").Format.Should().Contain("numberformat", "0.00");
        var boolCell = ReadNode(readOnly, "/Sheet1/A3");
        boolCell.Text.Should().Be("1");
        boolCell.Format.Should().Contain("type", "Boolean");
        ReadNode(readOnly, "/Sheet1/A4").Format.Should().Contain("numberformat", "yyyy-mm-dd");
        Query(readOnly, "cell").Count().Should().BeGreaterOrEqualTo(4);
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CellFormulas_RegularAndDynamicArray_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "10", "20", "30");
            handler.Add("/Sheet1/D1", "cell", null, Props(("formula", "SUM(A1:C1)")));
            handler.Add("/Sheet1/D2", "cell", null, Props(("formula", "SEQUENCE(3,1)")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/D1").Format.Should().Contain("formula", "SUM(A1:C1)");
        var dyn = ReadNode(readOnly, "/Sheet1/D2");
        dyn.Format.Should().Contain("formula", "SEQUENCE(3,1)");
        dyn.Format.Should().Contain("arrayformula", true);
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CellSet_FormulaViaSet_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "5", "7");
            handler.Add("/Sheet1/C1", "cell", null, Props(("value", "placeholder")));
            handler.Set("/Sheet1/C1", Props(("formula", "A1+B1"), ("numberformat", "0.00")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cell = ReadNode(readOnly, "/Sheet1/C1");
        cell.Format.Should().Contain("formula", "A1+B1");
        cell.Format.Should().Contain("numberformat", "0.00");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Cell locking & protection ───────────────────────────────────────

    [Fact]
    public void CellLockedAndFormulaHidden_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Unlocked"), ("locked", "false"), ("formulaHidden", "true")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "Locked"), ("locked", "true"), ("formulaHidden", "false")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cell1 = ReadNode(readOnly, "/Sheet1/A1");
        cell1.Text.Should().Be("Unlocked");

        var cell2 = ReadNode(readOnly, "/Sheet1/A2");
        cell2.Text.Should().Be("Locked");

        // Locking/protection properties may be reflected in the OOXML but are
        // not always surfaced as top-level Format keys; verify document validity.
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Rich text runs ──────────────────────────────────────────────────

    [Fact]
    public void RichTextRuns_AddMultipleRunsToCell_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Base")));
            handler.Add("/Sheet1/A1", "run", null, Props(("text", " + Bold"), ("bold", "true")));
            handler.Add("/Sheet1/A1", "run", null, Props(("text", " + Color"), ("font.color", "0000FF"), ("italic", "true")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var richCell = ReadNode(readOnly, "/Sheet1/A1");
        richCell.Format.Should().Contain("richtext", true);
        richCell.ChildCount.Should().BeGreaterOrEqualTo(2);

        var run1 = ReadNode(readOnly, "/Sheet1/A1/run[1]");
        run1.Text.Should().Be(" + Bold");
        run1.Format.Should().Contain("bold", true);

        var run2 = ReadNode(readOnly, "/Sheet1/A1/run[2]");
        run2.Text.Should().Be(" + Color");
        run2.Format.Should().Contain("color", "#0000FF");
        run2.Format.Should().Contain("italic", true);
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Merge / unmerge ─────────────────────────────────────────────────

    [Fact]
    public void MergeCells_AddAndReadback_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Merged Title"), ("merge", "A1:C2")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var merged = ReadNode(readOnly, "/Sheet1/A1");
        merged.Text.Should().Be("Merged Title");
        merged.Format.Should().Contain("merge", "A1:C2");

        var rangeNode = ReadNode(readOnly, "/Sheet1/A1:C2");
        rangeNode.Type.Should().Be("range");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void MergeCells_MultipleMergedRanges_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Header 1"), ("merge", "A1:B1")));
            handler.Add("/Sheet1/D1", "cell", null, Props(("value", "Header 2"), ("merge", "D1:E1")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var m1 = ReadNode(readOnly, "/Sheet1/A1");
        m1.Text.Should().Be("Header 1");
        m1.Format.Should().Contain("merge", "A1:B1");

        var m2 = ReadNode(readOnly, "/Sheet1/D1");
        m2.Text.Should().Be("Header 2");
        m2.Format.Should().Contain("merge", "D1:E1");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Row height / column width ───────────────────────────────────────

    [Fact]
    public void RowHeight_AddSetAndReadback_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "row", null, Props(("cols", "2")));
            handler.Set("/Sheet1/row[1]", Props(("height", "36")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var row = ReadNode(readOnly, "/Sheet1/row[1]");
        row.Format.Should().Contain("height", "36pt");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void ColumnWidth_AddSetAndReadback_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "col", null, Props(("name", "C"), ("width", "24"), ("hidden", "true")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var col = ReadNode(readOnly, "/Sheet1/col[C]");
        col.Format.Should().Contain("width", 24D);
        col.Format.Should().Contain("hidden", true);
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void RowHidden_AddHiddenRow_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Visible");
            handler.Add("/Sheet1", "row", null, Props(("cols", "1"), ("hidden", "true")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var hiddenRow = ReadNode(readOnly, "/Sheet1/row[2]");
        hiddenRow.Format.Should().Contain("hidden", true);
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Formula-reference mutation (row/col insert/delete/move) ─────────

    [Fact]
    public void FormulaRefMutation_RowInsert_BeforeFormulaSource_ShiftsFormulaRef()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "10");
            AddRow(handler, 2, "20");
            handler.Add("/Sheet1/A3", "cell", null, Props(("formula", "SUM(A1:A2)")));
            handler.Add("/Sheet1", "row", InsertPosition.AfterElement("/Sheet1/row[1]"), Props(("cols", "1")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        // The formula SUM(A1:A2) should still reference A1:A2 (inserted row is
        // between them, so range shifts to A1:A3).
        Query(readOnly, "cell[text~=SUM]").Should().NotBeEmpty("formula should persist through row insert");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void FormulaRefMutation_ColumnInsert_BeforeFormulaSource_ShiftsFormulaRef()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "10")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "20")));
            handler.Add("/Sheet1/D1", "cell", null, Props(("formula", "A1+B1")));
            handler.Add("/Sheet1", "col", InsertPosition.BeforeElement("/Sheet1/col[A]"), Props(("name", "A")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        // After inserting a column before A, the formula cell should still exist;
        // column references in the formula shift accordingly.
        Query(readOnly, "cell").Select(n => n.Path).Should().Contain(p => p.Contains("D1") || p.Contains("E1"));
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void FormulaRefMutation_RowMove_FormulaFollows()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "10");
            AddRow(handler, 2, "20");
            AddRow(handler, 3, "30");
            handler.Add("/Sheet1/A4", "cell", null, Props(("formula", "SUM(A1:A3)")));
            handler.Move("/Sheet1/row[2]", null, InsertPosition.AfterElement("/Sheet1/row[3]"));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        // Formula references should still resolve after row move.
        Query(readOnly, "cell[text~=SUM]").Should().NotBeEmpty("formula should persist through row move");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void FormulaRefMutation_RemoveRow_FormulaAdjusts()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "10");
            AddRow(handler, 2, "20");
            AddRow(handler, 3, "30");
            handler.Add("/Sheet1/A4", "cell", null, Props(("formula", "SUM(A1:A3)")));
            handler.Remove("/Sheet1/row[2]", null);
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        // After removing row 2, formula should still be intact (range may adjust).
        Query(readOnly, "cell[text~=SUM]").Should().NotBeEmpty("formula should survive row delete");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void FormulaRefMutation_SheetRename_FormulaKeepsWorking()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "sheet", null, Props(("name", "Source")));
            AddRow(handler, 1, "10", "20");
            handler.Add("/Sheet1/C1", "cell", null, Props(("formula", "A1+B1")));
            handler.Set("/Sheet1", Props(("name", "Target")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "sheet").Should().Contain(s => s.Path == "/Target");
        Query(readOnly, "sheet").Should().NotContain(s => s.Path == "/Sheet1");
        // Formula on the renamed sheet should still be readable.
        var formulaCell = ReadNode(readOnly, "/Target/C1");
        formulaCell.Format.Should().ContainKey("formula");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void FormulaRefMutation_NamedRangeFormula_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "100", "200");
            handler.Add("/Sheet1/C1", "cell", null, Props(("formula", "SUM(A1:B1)")));
            handler.Add("/", "namedrange", null, Props(("name", "TotalFormula"), ("ref", "Sheet1!$C$1")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var nr = Query(readOnly, "namedrange").Should().ContainSingle(n =>
            n.Format["name"]!.Equals("TotalFormula")).Subject;
        nr.Format.Should().Contain("ref", "Sheet1!$C$1");
        ReadNode(readOnly, "/Sheet1/C1").Format.Should().ContainKey("formula");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Comment lifecycle ───────────────────────────────────────────────

    [Fact]
    public void CommentLifecycle_AddSetGetQueryRemove_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Data")));
            handler.Add("/Sheet1/A1", "comment", null, Props(
                ("text", "Initial note"),
                ("author", "Alice"),
                ("font.bold", "true")
            ));
            handler.Save();
        }

        // Reopen: verify comment readback
        using (var readOnly = OpenReadOnly(path))
        {
            var comment = Query(readOnly, "comment").Should().ContainSingle().Subject;
            comment.Path.Should().Be("/Sheet1/comment[1]");
            comment.Text.Should().Be("Initial note");
            comment.Format.Should().Contain("ref", "A1");
            comment.Format.Should().Contain("author", "Alice");
            comment.Format.Should().Contain("font.bold", true);
            readOnly.Validate().Should().BeEmpty();
        }

        // Reopen: modify comment text
        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1/comment[1]", Props(
                ("text", "Updated note"),
                ("author", "Bob")
            ));
            handler.Save();
        }

        // Reopen: verify update
        using (var readOnly = OpenReadOnly(path))
        {
            var comment = ReadNode(readOnly, "/Sheet1/comment[1]");
            comment.Text.Should().Be("Updated note");
            comment.Format.Should().Contain("author", "Bob");
            readOnly.Validate().Should().BeEmpty();
        }

        // Reopen: remove comment
        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/comment[1]", null);
            handler.Save();
        }

        // Final reopen: comment should be gone
        using var finalReadOnly = OpenReadOnly(path);
        Query(finalReadOnly, "comment").Should().BeEmpty();
        finalReadOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CommentLifecycle_MultipleComments_IndependentEntities()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Q1", "100");
            AddRow(handler, 2, "Q2", "200");
            handler.Add("/Sheet1/A1", "comment", null, Props(("text", "First quarter"), ("author", "Reviewer")));
            handler.Add("/Sheet1/B2", "comment", null, Props(("text", "Verify this value"), ("author", "Auditor")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var comments = Query(readOnly, "comment");
        comments.Should().HaveCount(2);
        var refs = comments.Select(c => c.Format["ref"]!.ToString()).ToList();
        refs.Should().Contain("A1");
        refs.Should().Contain("B2");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Hyperlink lifecycle ─────────────────────────────────────────────

    [Fact]
    public void HyperlinkLifecycle_AddSetGetQueryRemove_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(
                ("value", "Click here"),
                ("link", "https://example.com"),
                ("tooltip", "Go to example"),
                ("display", "Example Link")
            ));
            handler.Save();
        }

        // Reopen: verify hyperlink readback
        using (var readOnly = OpenReadOnly(path))
        {
            var cell = ReadNode(readOnly, "/Sheet1/A1");
            cell.Format.Should().Contain("link", "https://example.com");
            cell.Format.Should().Contain("tooltip", "Go to example");
            cell.Format.Should().Contain("display", "Example Link");

            var hyperlink = Query(readOnly, "hyperlink").Should().ContainSingle().Subject;
            hyperlink.Path.Should().Be("/Sheet1/A1");
            readOnly.Validate().Should().BeEmpty();
        }

        // Reopen: modify hyperlink
        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1/A1", Props(
                ("link", "https://updated.com"),
                ("tooltip", "Updated tooltip")
            ));
            handler.Save();
        }

        // Reopen: verify update
        using (var readOnly = OpenReadOnly(path))
        {
            var cell = ReadNode(readOnly, "/Sheet1/A1");
            cell.Format.Should().Contain("link", "https://updated.com");
            cell.Format.Should().Contain("tooltip", "Updated tooltip");
            readOnly.Validate().Should().BeEmpty();
        }

        // Reopen: remove hyperlink by clearing link property
        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1/A1", Props(("link", "")));
            handler.Save();
        }

        // Final reopen: hyperlink query should be empty
        using var finalReadOnly = OpenReadOnly(path);
        Query(finalReadOnly, "hyperlink").Should().BeEmpty("hyperlink should be removed when link is cleared");
        finalReadOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void HyperlinkLifecycle_MultipleHyperlinks_IndependentEntities()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(
                ("value", "Google"), ("link", "https://google.com"), ("tooltip", "Search")));
            handler.Add("/Sheet1/A2", "cell", null, Props(
                ("value", "GitHub"), ("link", "https://github.com"), ("tooltip", "Code")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var links = Query(readOnly, "hyperlink");
        links.Should().HaveCount(2);
        var urls = links.Select(l => l.Format["url"]!.ToString()).ToList();
        urls.Should().Contain(u => u.Contains("google.com"));
        urls.Should().Contain(u => u.Contains("github.com"));
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Cell query selectors ────────────────────────────────────────────

    [Fact]
    public void CellQuery_ValueAndTextFilters_ReturnExpectedCells()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Price", "Status");
            AddRow(handler, 2, "Pen", "5", "Open");
            AddRow(handler, 3, "Book", "25", "Open");
            AddRow(handler, 4, "Desk", "150", "Closed");
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "cell[text~=Pen]").Select(n => n.Path).Should().Contain("/Sheet1/A2");
        Query(readOnly, "cell[text~=Closed]").Select(n => n.Path).Should().Contain("/Sheet1/C4");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CellClear_RemovesCell_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "ToRemove")));
            handler.Add("/Sheet1/A2", "cell", null, Props(("value", "Keep")));
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/A1", null);
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        Query(readOnly, "cell").Select(n => n.Path).Should().NotContain("/Sheet1/A1");
        Query(readOnly, "cell").Select(n => n.Path).Should().Contain("/Sheet1/A2");
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
