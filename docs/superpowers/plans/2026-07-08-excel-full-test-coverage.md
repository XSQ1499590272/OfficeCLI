# Excel Full Test Coverage Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add tracked tests that cover the full current Excel feature surface before Excel code slimming begins.

**Architecture:** Keep Excel tests isolated under `tests/OfficeCli.Tests/Excel/`. Use small direct handler tests for exact DOM/readback behavior and CLI smoke tests only where command routing, resident/batch, rendering, or file output must be exercised. Maintain `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md` as the checklist that maps every supported xlsx schema/example area to one or more tests.

**Tech Stack:** .NET 10, xUnit, FluentAssertions, DocumentFormat.OpenXml 3.4.1, existing `OfficeCli.Handlers.ExcelHandler`, `BlankDocCreator`, and CLI command surface.

## Global Constraints

- `tests/OfficeCli.Tests` is tracked. Do not add `tests/` back to `.gitignore`.
- Do not modify production code while adding Excel tests unless a test exposes an existing bug and the user explicitly approves the fix.
- Preserve Word, Excel, and PowerPoint render-look-fix behavior.
- Cover every current Excel public feature area represented by `schemas/help/xlsx/*.json`, `README.md`, `SKILL.md`, and `examples/excel/**`.
- Formula tests cover all advertised formula families and representative functions, not every one of the 350+ functions unless separately requested.
- If `dotnet --info` is unavailable, write the tests but report that execution is blocked until .NET 10 SDK is on `PATH`.

---

## File Structure

- Create: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`
  - Human-readable coverage manifest. Every xlsx schema key and every `examples/excel` category must map to a test class.
- Create: `tests/OfficeCli.Tests/Excel/ExcelTestBase.cs`
  - Shared temp file helpers and workbook open helpers.
- Create: `tests/OfficeCli.Tests/Excel/ExcelCellsAndFormulasTests.cs`
  - Workbook, sheet, row, column, cell, range, style, hyperlink, comments, formulas.
- Create: `tests/OfficeCli.Tests/Excel/ExcelTablesFiltersValidationTests.cs`
  - Tables/listobjects/detected tables, sorting, autofilter, validation, named ranges, data/query selectors.
- Create: `tests/OfficeCli.Tests/Excel/ExcelConditionalFormattingTests.cs`
  - Conditional formatting rule families: cellis, colorscale, databar, iconset, formulacf, topn, aboveaverage, duplicatevalues, uniquevalues, containstext, dateoccurring, cfextended.
- Create: `tests/OfficeCli.Tests/Excel/ExcelChartsAndDrawingsTests.cs`
  - Charts, chart axes/series, pictures, shapes, OLE, sparklines, row/column/page breaks.
- Create: `tests/OfficeCli.Tests/Excel/ExcelPivotAndSlicerTests.cs`
  - Pivot tables, pivot readback, slicers, pivot cache behavior smoke.
- Create: `tests/OfficeCli.Tests/Excel/ExcelDumpBatchRenderTests.cs`
  - Dump/batch round-trip, raw access, validate, view html, view screenshot where practical.

### Task 1: Establish Coverage Manifest And Shared Helpers

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`
- Create: `tests/OfficeCli.Tests/Excel/ExcelTestBase.cs`

**Interfaces:**
- Produces: `ExcelTestBase.CreateWorkbook()` returning a temp `.xlsx` path.
- Produces: `ExcelTestBase.OpenEditable(string path)` returning `ExcelHandler`.
- Produces: `ExcelTestBase.ReadNode(ExcelHandler handler, string path)` returning non-null `DocumentNode`.

- [ ] **Step 1: Create the coverage manifest**

Create `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md` with this exact coverage table:

```markdown
# Excel Test Coverage

Every checked item must have at least one tracked test before Excel code slimming starts.

## Schema Surface

- [ ] aboveaverage -> ExcelConditionalFormattingTests
- [ ] autofilter -> ExcelTablesFiltersValidationTests
- [ ] cell -> ExcelCellsAndFormulasTests
- [ ] cellis -> ExcelConditionalFormattingTests
- [ ] cfextended -> ExcelConditionalFormattingTests
- [ ] chart -> ExcelChartsAndDrawingsTests
- [ ] chart-axis -> ExcelChartsAndDrawingsTests
- [ ] chart-series -> ExcelChartsAndDrawingsTests
- [ ] colbreak -> ExcelChartsAndDrawingsTests
- [ ] colorscale -> ExcelConditionalFormattingTests
- [ ] column -> ExcelCellsAndFormulasTests
- [ ] comment -> ExcelCellsAndFormulasTests
- [ ] conditionalformatting -> ExcelConditionalFormattingTests
- [ ] containstext -> ExcelConditionalFormattingTests
- [ ] databar -> ExcelConditionalFormattingTests
- [ ] dateoccurring -> ExcelConditionalFormattingTests
- [ ] detectedtable -> ExcelTablesFiltersValidationTests
- [ ] duplicatevalues -> ExcelConditionalFormattingTests
- [ ] formulacf -> ExcelConditionalFormattingTests
- [ ] hyperlink -> ExcelCellsAndFormulasTests
- [ ] iconset -> ExcelConditionalFormattingTests
- [ ] namedrange -> ExcelTablesFiltersValidationTests
- [ ] ole -> ExcelChartsAndDrawingsTests
- [ ] pagebreak -> ExcelChartsAndDrawingsTests
- [ ] picture -> ExcelChartsAndDrawingsTests
- [ ] pivottable -> ExcelPivotAndSlicerTests
- [ ] range -> ExcelCellsAndFormulasTests
- [ ] raw -> ExcelDumpBatchRenderTests
- [ ] row -> ExcelCellsAndFormulasTests
- [ ] rowbreak -> ExcelChartsAndDrawingsTests
- [ ] run -> ExcelCellsAndFormulasTests
- [ ] shape -> ExcelChartsAndDrawingsTests
- [ ] sheet -> ExcelCellsAndFormulasTests
- [ ] slicer -> ExcelPivotAndSlicerTests
- [ ] sort -> ExcelTablesFiltersValidationTests
- [ ] sparkline -> ExcelChartsAndDrawingsTests
- [ ] table -> ExcelTablesFiltersValidationTests
- [ ] topn -> ExcelConditionalFormattingTests
- [ ] uniquevalues -> ExcelConditionalFormattingTests
- [ ] validation -> ExcelTablesFiltersValidationTests
- [ ] workbook -> ExcelCellsAndFormulasTests

## Example Surface

- [ ] cell-formatting -> ExcelCellsAndFormulasTests
- [ ] charts/* -> ExcelChartsAndDrawingsTests
- [ ] conditional-formatting -> ExcelConditionalFormattingTests
- [ ] data-validation -> ExcelTablesFiltersValidationTests
- [ ] pivot-tables -> ExcelPivotAndSlicerTests
- [ ] shapes -> ExcelChartsAndDrawingsTests
- [ ] sheet-settings -> ExcelCellsAndFormulasTests
- [ ] slicers -> ExcelPivotAndSlicerTests
- [ ] sparklines -> ExcelChartsAndDrawingsTests
- [ ] workbook-settings -> ExcelCellsAndFormulasTests

## Command Surface

- [ ] create -> ExcelCellsAndFormulasTests
- [ ] add -> all Excel test classes
- [ ] get -> all Excel test classes
- [ ] query -> all Excel test classes
- [ ] set -> all Excel test classes
- [ ] remove -> ExcelCellsAndFormulasTests, ExcelTablesFiltersValidationTests
- [ ] move -> ExcelCellsAndFormulasTests
- [ ] batch -> ExcelDumpBatchRenderTests
- [ ] dump -> ExcelDumpBatchRenderTests
- [ ] raw -> ExcelDumpBatchRenderTests
- [ ] validate -> ExcelDumpBatchRenderTests
- [ ] view html -> ExcelDumpBatchRenderTests
- [ ] view screenshot -> ExcelDumpBatchRenderTests
```

- [ ] **Step 2: Create shared test helpers**

Create `tests/OfficeCli.Tests/Excel/ExcelTestBase.cs`:

```csharp
// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Excel;

public abstract class ExcelTestBase : IDisposable
{
    private readonly List<string> _paths = new();

    protected string CreateWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_excel_{Guid.NewGuid():N}.xlsx");
        _paths.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    protected ExcelHandler OpenEditable(string path) => new(path, editable: true);

    protected ExcelHandler OpenReadOnly(string path) => new(path, editable: false);

    protected static DocumentNode ReadNode(ExcelHandler handler, string path)
    {
        var node = handler.Get(path);
        node.Should().NotBeNull(path);
        return node!;
    }

    protected static IReadOnlyList<DocumentNode> Query(ExcelHandler handler, string selector)
        => handler.Query(selector);

    public void Dispose()
    {
        foreach (var path in _paths)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }
}
```

- [ ] **Step 3: Run compile check**

Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --no-restore
```

Expected if .NET 10 SDK and packages are available:

```text
Passed!
```

If `dotnet` is missing, record the blocker and continue writing tests without claiming they pass.

### Task 2: Cells, Sheets, Rows, Columns, Ranges, Formulas

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelCellsAndFormulasTests.cs`
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Interfaces:**
- Consumes: `ExcelTestBase`.
- Produces: tests covering `workbook`, `sheet`, `row`, `column`, `cell`, `range`, `run`, `comment`, `hyperlink`, formula families, sheet settings, workbook settings.

- [ ] **Step 1: Add tests for sheet/cell/row/column/range basics**

Create tests that:

- create a workbook
- set `/Sheet1/A1` text, number, date, boolean, formula, rich text
- get `/Sheet1/A1`, `/Sheet1/row[1]`, `/Sheet1/col[A]`, and `/Sheet1/A1:B2`
- add a sheet with visible/hidden/veryHidden states
- set row height and column width
- merge a range and assert merged readback

- [ ] **Step 2: Add formula family tests**

Cover representative formulas:

- arithmetic: `=SUM(A1:A3)`
- lookup: `=XLOOKUP(...)` or `=VLOOKUP(...)`
- dynamic array: `=SEQUENCE(2,2)`
- text/date: `=TEXT(...)`, `=DATE(...)`
- financial/bond/statistical: one representative from each advertised family
- defined-name formula body inlining
- formula reference rewrite after row/column insert

- [ ] **Step 3: Add comments, hyperlinks, and workbook/sheet setting tests**

Cover:

- cell comment add/get/query, including RTL flag
- hyperlink add/get/query
- workbook password/properties where supported
- sheet print settings, RTL sheet view, print titles, margins

- [ ] **Step 4: Mark coverage items complete**

Check the corresponding `EXCEL_COVERAGE.md` items only after tests exist.

### Task 3: Tables, Filters, Validation, Named Ranges, Sort

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelTablesFiltersValidationTests.cs`
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Interfaces:**
- Consumes: `ExcelTestBase`.
- Produces: tests covering `table`, `detectedtable`, `autofilter`, `sort`, `validation`, `namedrange`, row-by-column selectors.

- [ ] **Step 1: Add table/listobject and detected table tests**

Cover table add/query/get/set/remove, style readback, and detected table query for plain data blocks.

- [ ] **Step 2: Add autofilter and sort tests**

Cover sheet/range sort, multi-key sort, autofilter add/query, and row selector predicates such as `row[Salary>5000 and Region=EMEA]`.

- [ ] **Step 3: Add data validation and named range tests**

Cover list/whole/decimal/date/textLength/custom validations, named ranges, defined names, and formula-body readback.

- [ ] **Step 4: Mark coverage items complete**

Check the corresponding `EXCEL_COVERAGE.md` items only after tests exist.

### Task 4: Conditional Formatting

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelConditionalFormattingTests.cs`
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Interfaces:**
- Consumes: `ExcelTestBase`.
- Produces: tests covering all conditional formatting schema families.

- [ ] **Step 1: Add conditional formatting tests by rule family**

Cover:

- `cellis`
- `colorscale`
- `databar`
- `iconset`
- `formulacf`
- `topn`
- `aboveaverage`
- `duplicatevalues`
- `uniquevalues`
- `containstext`
- `dateoccurring`
- `cfextended`

- [ ] **Step 2: Verify query/get/readback**

Each rule family must be queryable and must preserve its key properties after save and reopen.

- [ ] **Step 3: Mark coverage items complete**

Check the corresponding `EXCEL_COVERAGE.md` items only after tests exist.

### Task 5: Charts, Drawings, Pictures, Shapes, OLE, Sparklines, Breaks

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelChartsAndDrawingsTests.cs`
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Interfaces:**
- Consumes: `ExcelTestBase`.
- Produces: tests covering chart families, axes, series, drawing objects, pictures, shapes, OLE, sparklines, row/column/page breaks.

- [ ] **Step 1: Add chart tests**

Cover standard chart types and extended chart families represented in examples:

- basic
- area
- bar
- column
- line
- pie
- combo
- scatter
- bubble
- radar
- stock
- waterfall
- histogram
- boxwhisker
- advanced properties such as log axis, title, legend, anchor, series query, axis query

- [ ] **Step 2: Add drawing object tests**

Cover picture add/query/get/set, SVG picture fallback, shape add/query/get/set, and OLE object query/readback.

- [ ] **Step 3: Add sparklines and breaks tests**

Cover sparkline add/query/get and rowbreak/colbreak/pagebreak readback.

- [ ] **Step 4: Mark coverage items complete**

Check the corresponding `EXCEL_COVERAGE.md` items only after tests exist.

### Task 6: Pivot Tables And Slicers

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelPivotAndSlicerTests.cs`
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Interfaces:**
- Consumes: `ExcelTestBase`.
- Produces: tests covering pivot tables, pivot cache/readback, and slicers.

- [ ] **Step 1: Add pivot table tests**

Cover:

- row/column/filter/value fields
- aggregation functions
- grand totals
- subtotals
- compact/outline/tabular layout
- repeat item labels
- blank rows
- date grouping
- calculated fields
- showDataAs
- topN and labelFilter
- cache copy-on-write and cross-pivot sharing smoke

- [ ] **Step 2: Add slicer tests**

Cover slicer add/query/get/readback against a table-backed source and a pivot-backed source if supported.

- [ ] **Step 3: Mark coverage items complete**

Check the corresponding `EXCEL_COVERAGE.md` items only after tests exist.

### Task 7: Dump, Batch, Raw, Render, Validate

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelDumpBatchRenderTests.cs`
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Interfaces:**
- Consumes: `ExcelTestBase`.
- Produces: tests covering dump/batch replay, raw XML paths, validation, HTML rendering, screenshot rendering where environment supports it.

- [ ] **Step 1: Add dump and batch round-trip tests**

Cover whole workbook dump, sheet subtree dump, table/chart/pivot/slicer carrier behavior where supported, and replay through batch.

- [ ] **Step 2: Add raw path tests**

Cover `/workbook`, `/styles`, `/sharedstrings`, `/theme`, `/<SheetName>`, `/<SheetName>/drawing`, `/<SheetName>/chart[N]`, and zip-internal XML paths.

- [ ] **Step 3: Add validate and issue tests**

Cover `Validate()` returning no hard schema errors for generated workbooks and `view issues` behavior through the command layer if practical.

- [ ] **Step 4: Add render tests**

Cover `ViewAsHtml()` directly and CLI `view html`. Cover screenshot through CLI when Chrome/headless screenshot support is available; otherwise assert the test is skipped with a clear reason.

- [ ] **Step 5: Mark coverage items complete**

Check the corresponding `EXCEL_COVERAGE.md` items only after tests exist.

### Task 8: Final Excel Gate

**Files:**
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Interfaces:**
- Consumes: every Excel test class.
- Produces: confirmed Excel test coverage gate for later slimming.

- [ ] **Step 1: Run all Excel tests**

Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~OfficeCli.Tests.Excel"
```

Expected:

```text
Passed!
```

- [ ] **Step 2: Run full tracked test project**

Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected:

```text
Passed!
```

- [ ] **Step 3: Confirm manifest has no unchecked items**

Run:

```bash
rg -n "^- \\[ \\]" tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md
```

Expected: no output.

- [ ] **Step 4: Commit only tests and coverage docs**

Run:

```bash
git add tests/OfficeCli.Tests/Excel docs/superpowers/plans/2026-07-08-excel-full-test-coverage.md
git commit -m "test(excel): cover full feature surface"
```
