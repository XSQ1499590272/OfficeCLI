# Excel Full Test Matrix Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Complete Unit, Integration, and E2E tests for every currently implemented or exposed Excel feature in OfficeCLI.

**Architecture:** Keep one existing xUnit project and add focused Excel test files under `tests/OfficeCli.Tests/Excel`. Reuse `ExcelTestBase` and current direct `ExcelHandler` patterns; use CLI process tests only where the process boundary, resident flow, examples, render output, or command routing is the behavior under test. Treat the current 45 passing tests as baseline coverage, then fill the matrix gaps from `docs/superpowers/specs/2026-07-08-excel-full-test-matrix-design.md`.

**Tech Stack:** .NET 10, xUnit, FluentAssertions, DocumentFormat.OpenXml, existing `OfficeCli.Handlers.ExcelHandler`, existing `CommandBuilder`, existing `BlankDocCreator`, existing `./build.sh all`.

## Global Constraints

- Scope is current implemented or exposed Excel behavior only; do not add new Excel product features.
- Do not add a second test project.
- Do not add new NuGet packages.
- Do not add automatic commit steps; the user owns final submission.
- Production-code edits are allowed only for test visibility or for a real bug exposed by tests after user approval.
- Each implementation batch must run its filtered tests and then the full `OfficeCli.Tests` suite.
- Run `./build.sh all` once after the full Excel matrix is complete.
- Screenshot E2E may accept `no_screenshot_backend`; `view html` must produce inspectable HTML.

---

## File Structure

- Modify: `tests/OfficeCli.Tests/Excel/ExcelTestBase.cs`
  - Shared helpers for temp directories, CLI execution, example script execution, zip part reads, small PNG/OLE/CSV fixtures.
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`
  - Replace broad checkmarks with a matrix that marks Unit, Integration, and E2E for each feature area.
- Create: `tests/OfficeCli.Tests/Excel/ExcelFormulaUnitTests.cs`
  - Formula evaluator, formula cache, modern qualifier, reference shifter.
- Create: `tests/OfficeCli.Tests/Excel/ExcelFormattingUnitTests.cs`
  - Number/date/text formatting, style keys, color helpers where accessible.
- Create: `tests/OfficeCli.Tests/Excel/ExcelSelectorsAndRawUnitTests.cs`
  - Attribute filters, selector aliases, raw XML actions, worksheet bloat filter.
- Create: `tests/OfficeCli.Tests/Excel/ExcelChartUnitTests.cs`
  - Chart range parsing, chart setters, renderer extraction helpers where accessible.
- Create: `tests/OfficeCli.Tests/Excel/ExcelPivotTableUnitTests.cs`
  - Pivot parser/cache/readback/set helpers where accessible.
- Create: `tests/OfficeCli.Tests/Excel/ExcelRenderUnitTests.cs`
  - HTML screenshot helpers and deterministic render helpers.
- Create: `tests/OfficeCli.Tests/Excel/ExcelWorkbookIntegrationTests.cs`
  - Workbook lifecycle, workbook settings, sheet settings, sheet mutation.
- Create: `tests/OfficeCli.Tests/Excel/ExcelCellsRangesIntegrationTests.cs`
  - Cells, ranges, rich text, formulas, comments, hyperlinks, row/column/range lifecycle.
- Create: `tests/OfficeCli.Tests/Excel/ExcelTablesFiltersValidationIntegrationTests.cs`
  - Tables, detected tables, filters, sorting, validation, named ranges.
- Create: `tests/OfficeCli.Tests/Excel/ExcelConditionalFormattingIntegrationTests.cs`
  - Every implemented conditional formatting family.
- Create: `tests/OfficeCli.Tests/Excel/ExcelChartsDrawingsIntegrationTests.cs`
  - Standard charts, extended charts, pictures, shapes, OLE, sparklines.
- Create: `tests/OfficeCli.Tests/Excel/ExcelPivotSlicerIntegrationTests.cs`
  - Pivot tables and slicers.
- Create: `tests/OfficeCli.Tests/Excel/ExcelRawDumpImportIntegrationTests.cs`
  - Raw/raw-set, dump/batch replay, import, bloat filter.
- Create: `tests/OfficeCli.Tests/Excel/ExcelCliCommandE2ETests.cs`
  - Real CLI coverage for every Excel command in the design matrix.
- Create: `tests/OfficeCli.Tests/Excel/ExcelExamplesE2ETests.cs`
  - Real CLI coverage for every deterministic `examples/excel/**/*.sh` script.
- Create: `tests/OfficeCli.Tests/Excel/ExcelRenderE2ETests.cs`
  - Real CLI coverage for `view html` and `view screenshot`.

## Task 1: Coverage Manifest And Shared Test Helpers

**Files:**
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`
- Modify: `tests/OfficeCli.Tests/Excel/ExcelTestBase.cs`

**Interfaces:**
- Produces: `TrackTempDirectory() : string`
- Produces: `ReadZipEntry(string workbookPath, string entryName) : string`
- Produces: `CopyExampleScriptToTemp(string relativeScriptPath) : string`
- Produces: `RunShellScriptOk(string scriptPath, string workingDirectory, IReadOnlyDictionary<string,string>? env = null) : CliResult`
- Produces: `CreateTinyPng(string path) : string`
- Produces: `CreateOlePayload(string path) : string`
- Produces: `WriteCsv(string path, params string[] rows) : string`

- [x] Add helper methods above to `ExcelTestBase`.
- [x] Update `EXCEL_COVERAGE.md` so each matrix row has `Unit`, `Integration`, and `E2E` columns.
- [x] Mark existing coverage only where a current test already proves that layer.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelTestBase|FullyQualifiedName~ExcelUnitTests|FullyQualifiedName~ExcelIntegrationTests|FullyQualifiedName~ExcelE2ETests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: filtered tests pass and the full suite passes.

## Task 2: Full Formula And Reference Unit Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelFormulaUnitTests.cs`

**Interfaces:**
- Consumes: `src/officecli/Core/Formula/FormulaEvaluator*.cs`
- Consumes: `src/officecli/Core/Formula/ModernFunctionQualifier.cs`
- Consumes: `src/officecli/Core/FormulaRefShifter.cs`
- Consumes: `src/officecli/Handlers/Excel/ExcelHandler.FormulaCache.cs`

- [x] Add `FormulaEvaluator_CoversAllImplementedDispatchFamilies`.
  - Use theory data grouped by the current evaluator files: arithmetic/references, lookup, text/date/time, dynamic array/spill, financial, statistics, regression, securities, special functions, error functions.
  - Include every implemented function family currently dispatched in `FormulaEvaluator.*`; assertions must cover success values and status.
- [x] Add `FormulaEvaluator_PropagatesExcelErrorsAndUnsupportedFunctions`.
  - Cover `#DIV/0!`, `#NUM!`, `#VALUE!`, missing refs, bad argument counts, unknown functions, blank operands.
- [x] Add `ModernFunctionQualifier_QualifiesDynamicArraysAndQuotesSheetRefs`.
  - Cover dynamic array detection, `_xlfn` qualification, unqualification, sheet names with spaces/punctuation.
- [x] Add `FormulaRefShifter_RewritesRowsColumnsCopiesAndSheetRenames`.
  - Cover absolute refs, relative refs, ranges, cross-sheet refs, quoted sheet names, structured text that must not be rewritten.
- [x] Add `FormulaCacheHelpers_ClassifyAllowlistAndComputedAgreement`.
  - Cover allowlist hits/misses, numeric tolerance, string/error mismatches.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelFormulaUnitTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: formula unit tests pass and the full suite passes.

## Task 3: Formatting, Selectors, Raw, Bloat, And Render Unit Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelFormattingUnitTests.cs`
- Create: `tests/OfficeCli.Tests/Excel/ExcelSelectorsAndRawUnitTests.cs`
- Create: `tests/OfficeCli.Tests/Excel/ExcelRenderUnitTests.cs`

**Interfaces:**
- Consumes: `ExcelDataFormatter`
- Consumes: `AttributeFilter`
- Consumes: `RawXmlHelper`
- Consumes: `WorksheetBloatFilter`
- Consumes: `HtmlScreenshot`

- [x] Add formatting tests for every built-in format family used by Excel readback: general, integer, decimal, percent, scientific, fraction, currency/accounting, date, time, text, custom positive/negative/zero/text sections.
- [x] Add selector tests for equality, inequality, contains, numeric comparison, exists, regex find, `and`, `or`, unknown-key diagnostics, Excel cell alias normalization.
- [x] Add raw XML tests for `append`, `prepend`, `insertbefore`, `insertafter`, `replace`, `remove`, `setattr`, namespace preservation, schema-order insertion, zip URI path detection.
- [x] Add worksheet bloat tests for filtering empty high-index cells while preserving cells with value, formula, or style.
- [x] Add render helper tests for screenshot grid sizing, backend detection result handling, color normalization, number-format section selection where accessible.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelFormattingUnitTests|FullyQualifiedName~ExcelSelectorsAndRawUnitTests|FullyQualifiedName~ExcelRenderUnitTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: unit tests pass and the full suite passes.

## Task 4: Chart And Pivot Unit Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelChartUnitTests.cs`
- Create: `tests/OfficeCli.Tests/Excel/ExcelPivotTableUnitTests.cs`

**Interfaces:**
- Consumes: `ChartHelper`
- Consumes: `ChartExBuilder`
- Consumes: `ChartSvgRenderer`
- Consumes: `PivotTableHelper`

- [x] Add chart helper tests for range parsing, series refs, chart type aliases, axis property normalization, title/legend/label setters, trendline/error-bar options, secondary axis mapping.
- [x] Add chartEx helper tests for extended chart resource loading and range remapping.
- [x] Add chart renderer extraction tests for standard chart data, per-point colors, deleted labels, marker settings, trendlines, error bars, and reference lines.
- [x] Add pivot helper tests for cache parsing, field/value/filter parsing, settable properties, refresh flags, and orphan cache pruning decisions where helperized.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelChartUnitTests|FullyQualifiedName~ExcelPivotTableUnitTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: chart/pivot unit tests pass and the full suite passes.

## Task 5: Workbook, Sheet, Cell, Range, And Formula Integration Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelWorkbookIntegrationTests.cs`
- Create: `tests/OfficeCli.Tests/Excel/ExcelCellsRangesIntegrationTests.cs`

**Interfaces:**
- Consumes: `ExcelTestBase.OpenEditable`
- Consumes: `ExcelHandler.Add`, `Set`, `Get`, `Query`, `Remove`, `Move`, `Save`

- [x] Add workbook lifecycle tests for create/open/save/reopen/validate, workbook properties, calc settings, protection, active tab, first sheet.
- [x] Add sheet lifecycle tests for add/rename/visibility/freeze/gridlines/RTL/tab color/protection/print area/print titles/margins/header/footer/page setup.
- [x] Add cell/range tests for text, number, boolean, date, formula, array formula, clear, type, locked, formula hidden, rich text runs, merge/unmerge, row height, column width.
- [x] Add formula reference mutation tests for row/column insert/delete/move/copy, sheet rename, named range formulas, chart/table refs where currently rewritten.
- [x] Add comment and hyperlink lifecycle tests with add/set/get/query/remove and save/reopen.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelWorkbookIntegrationTests|FullyQualifiedName~ExcelCellsRangesIntegrationTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: workbook/cell integration tests pass and the full suite passes.

## Task 6: Tables, Filters, Validation, Named Ranges, And Conditional Formatting Integration Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelTablesFiltersValidationIntegrationTests.cs`
- Create: `tests/OfficeCli.Tests/Excel/ExcelConditionalFormattingIntegrationTests.cs`

**Interfaces:**
- Consumes: `ExcelHandler.Add`, `Set`, `Get`, `Query`, `Remove`, `Save`

- [x] Add table tests for add/get/query/set/remove, table style, totals row, column rename, auto expansion, detected table readback.
- [x] Add autofilter tests for equals, notEquals, contains, notContains, beginsWith, endsWith, greater/less comparisons, between/notBetween, top/bottom, blanks/nonblanks, values, dynamic filters.
- [x] Add sort tests for single/multiple keys, header/no-header, numeric/text/date/null ordering.
- [x] Add validation tests for list, whole, decimal, date, time, textLength, custom, operators, allowBlank, prompt/error messages, dropdown behavior.
- [x] Add named range tests for workbook scope, sheet scope, comments, volatile flag, formula use, set/remove.
- [x] Add conditional formatting tests for cellis, colorscale, databar, iconset, formula, top/bottom, above/below average, unique, duplicate, contains text, contains blanks, not contains blanks, contains errors, not contains errors, begins/ends, date occurring, cfextended.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelTablesFiltersValidationIntegrationTests|FullyQualifiedName~ExcelConditionalFormattingIntegrationTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: table/filter/validation/CF integration tests pass and the full suite passes.

## Task 7: Charts, Drawings, Pivot, Slicer, Raw, Dump, Import Integration Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelChartsDrawingsIntegrationTests.cs`
- Create: `tests/OfficeCli.Tests/Excel/ExcelPivotSlicerIntegrationTests.cs`
- Create: `tests/OfficeCli.Tests/Excel/ExcelRawDumpImportIntegrationTests.cs`

**Interfaces:**
- Consumes: `ExcelHandler`
- Consumes: `ExcelBatchEmitter.EmitExcel`
- Consumes: `CommandBuilder.RunNonResidentBatch`

- [x] Add standard chart tests for column, bar, line, pie, area, scatter, bubble, radar, stock, combo, axes, series, labels, legend, title, layout, trendline, error bars, raw chart access, remove.
- [x] Add extended chart tests for waterfall, histogram, boxWhisker and other currently implemented chartEx paths.
- [x] Add drawing tests for png/jpg/svg pictures, alt text, crop, hyperlink, anchors, binary extraction, shapes, shape text/font/fill/line/margins/align/valign/gradient, OLE payloads, sparklines.
- [x] Add pivot tests for add/get/set/remove, cache readback, fields, values, filters, refresh flags, orphan cleanup.
- [x] Add slicer tests for add/get/set/remove and relationship readback.
- [x] Add raw/raw-set tests for semantic aliases, zip URI worksheet paths, row/column filters, workbook/sheet/chart/drawing mutation, validate after mutation.
- [x] Add dump/batch tests that replay emitted commands into a blank workbook and compare workbook, sheet, table, CF, chart, drawing, pivot, slicer, and raw essentials.
- [x] Add import tests for CSV/TSV into a workbook and query/readback.
- [x] Add bloat filter integration test using a generated workbook with huge empty cells and value/formula/style cells that must survive.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelChartsDrawingsIntegrationTests|FullyQualifiedName~ExcelPivotSlicerIntegrationTests|FullyQualifiedName~ExcelRawDumpImportIntegrationTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: chart/drawing/pivot/raw/import integration tests pass and the full suite passes.

## Task 8: Complete CLI Command E2E Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelCliCommandE2ETests.cs`

**Interfaces:**
- Consumes: `ExcelTestBase.RunCliOk`
- Consumes: `ExcelTestBase.RunCli`

- [x] Add real-process E2E tests for `create`, `open`, `save`, `close`, `add`, `get`, `query`, `set`, `remove`, `move`, `swap`, `batch`, `dump`, `raw`, `raw-set`, `validate`, `import`.
- [x] Cover text output and JSON output for `get`, `query`, errors, and warnings.
- [x] Cover non-zero CLI errors for missing file, invalid sheet/range/type/prop, unknown raw part, invalid formula, duplicate object where currently rejected.
- [x] Verify every mutating command leaves the workbook valid with `validate`.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelCliCommandE2ETests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: CLI command E2E tests pass and the full suite passes.

## Task 9: Complete Examples And Render E2E Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelExamplesE2ETests.cs`
- Create: `tests/OfficeCli.Tests/Excel/ExcelRenderE2ETests.cs`

**Interfaces:**
- Consumes: `ExcelTestBase.CopyExampleScriptToTemp`
- Consumes: `ExcelTestBase.RunShellScriptOk`
- Consumes: `ExcelTestBase.RunCliOk`

- [x] Add E2E coverage for every script listed in the design matrix under `examples/excel/**/*.sh`.
- [x] Run scripts in temp directories so tracked example workbooks are not modified.
- [x] Put a temp `officecli` wrapper at the front of `PATH` so scripts use the current test build.
- [x] Validate every generated workbook after the script exits.
- [x] Add `view html --out` tests for cells, formats, formulas, CF, charts, drawings, freeze panes, and overflow-sensitive text.
- [x] Add `view screenshot` tests that accept either a PNG file or `no_screenshot_backend`.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelExamplesE2ETests|FullyQualifiedName~ExcelRenderE2ETests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: examples/render E2E tests pass and the full suite passes.

## Task 10: Coverage Manifest Closure And Final Verification

**Files:**
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Interfaces:**
- Consumes: all test classes from Tasks 1-9.

- [x] Update every `EXCEL_COVERAGE.md` matrix row with the exact test class and method names that cover Unit, Integration, and E2E.
- [x] Leave `N/A` only where the design allows it and include a reason in the same row.
- [x] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
./build.sh all
```

Expected: full suite passes and all 8 release binaries build.

## Self-Review

- Spec coverage: every feature row from `2026-07-08-excel-full-test-matrix-design.md` maps to at least one task.
- CLI coverage: every command in the design CLI matrix maps to Task 8 or Task 9.
- Example coverage: every script in the design example matrix maps to Task 9.
- Verification coverage: every implementation batch runs filtered tests plus full tests; final batch runs `./build.sh all`.
- Scope control: no new packages, no new test project, no automatic commit steps.
