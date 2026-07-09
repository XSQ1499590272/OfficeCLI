# Excel Full Test Matrix Design

## Scope

Cover the Excel functionality that is currently implemented or exposed by this repository. Do not add new Excel features as part of this work.

The goal is not a small smoke suite. The goal is a regression net strong enough to support later Excel code slimming:

- Unit tests cover pure/internal logic and error branches.
- Integration tests cover `ExcelHandler` behavior with save/reopen readback.
- E2E tests cover real CLI process routing, resident flow, batch/dump/raw/view/import, and examples.
- Cross-platform packaging is run once after the full matrix is complete, not after each small test batch.

## Code-Fact Sources

- CLI commands are registered in `src/officecli/CommandBuilder*.cs`.
- Excel implementation lives under `src/officecli/Handlers/Excel*` and Excel-specific core helpers under `src/officecli/Core`.
- Existing Excel examples live under `examples/excel/**`.
- Current Excel tests live under `tests/OfficeCli.Tests/Excel`.
- `tests/OfficeCli.Tests` targets `net10.0` and references `src/officecli/officecli.csproj`.

## Completeness Rule

A feature is complete only when all applicable layers are present:

- Unit: deterministic helper/parser/formatter/validator logic, including invalid inputs.
- Integration: public handler API creates or mutates a workbook, saves, reopens, and verifies semantic readback or raw OOXML shape.
- E2E: CLI command or example script exercises the feature through the real process boundary.

If a feature has no meaningful unit layer because all logic is direct Open XML orchestration, the matrix must mark Unit as `N/A` with the reason.

## Feature Matrix

| Area | Current feature surface | Unit coverage target | Integration coverage target | E2E coverage target |
| --- | --- | --- | --- | --- |
| Workbook lifecycle | `create`, `open`, `save`, `close`, `validate`, workbook properties, calc settings, protection | workbook property parsing and validation helpers | create workbook, set workbook/calc/protection settings, save/reopen, validate | CLI create/open/batch/save/close/validate |
| Sheets | add sheet, rename, visibility, active tab, first sheet, freeze, gridlines, RTL, tab color, protection, print area/titles, margins, headers/footers, page setup | sheet name/range validation, sheet ref rename helpers | add/set/remove/move sheets and verify workbook refs after reopen | CLI add/set/get/remove/move sheet paths |
| Cells and ranges | value, formula, array formula, type, clear, rich text runs, style, number format, locked/formulaHidden, merge, range/row/column readback | cell address/range parsing, value length validation, style key mapping | add/set/get/query cells/ranges/rows/columns, merge/unmerge, save/reopen | CLI add/set/get/query on `Sheet!A1`, row, col, range |
| Formula engine | arithmetic, refs, ranges, lookups, text, date/time, dynamic arrays, financial, statistics, regression, securities, special functions, errors | one or more tests per implemented formula family plus bad args/error propagation | formulas written to workbook produce expected cached/readback values where evaluator supports them | CLI formula set/get through real workbook |
| Formula references | row/column insert/delete/move/copy, sheet rename, named range references | `FormulaRefShifter`, cache allowlist, computed/cache comparison | mutations rewrite formulas, named ranges, chart/table refs when applicable | CLI move/remove/set followed by get/raw |
| Formatting | built-in/custom number formats, dates, percent, currency, text, font, fill, border, alignment, hyperlinks, display value | `ExcelDataFormatter`, HTML number/date/color helpers | style and hyperlink round-trip through handler | CLI cell-formatting example plus direct set/get |
| Comments | add/set/get/query/remove comments, author, RTL, rich text | comment property validation helpers if exposed, otherwise N/A | comment add/set/remove survives save/reopen | CLI add/set/get/remove comment |
| Tables | table add/get/query/set/remove, style, totals, columns, auto expansion, detected tables | table helper mapping, totals function mapping, table style resolver | table lifecycle, detected table query, total row/columns after reopen | CLI table and detectedtable commands |
| Autofilter and sort | autofilter criteria families, sort keys, header handling | criteria parsing, sort value parsing, selector warnings | add/set/remove autofilter, sort rows, verify order/readback | CLI data-validation/table scripts plus direct sort |
| Data validation | list, whole, decimal, date, time, textLength, custom, operator, prompts/errors/dropdown | sqref/range validation and enum parsing | all validation families add/set/get/remove after reopen | CLI data-validation example and direct validation commands |
| Named ranges | workbook/sheet-scoped names, comments, volatile flag, ref updates | name/ref validation, sheet-scope parsing | add/set/get/query/remove named ranges, formula use | CLI namedrange add/set/get/remove |
| Conditional formatting | cellis, colorscale, databar, iconset, formula, top/bottom, above/below average, unique/duplicate, contains text/blanks/errors, date occurring, extended rules | operator/type parsing, color/icon/databar property parsing, HTML CF evaluation helpers | add/set/get/remove every implemented CF family after reopen | CLI conditional-formatting example and direct CF commands |
| Charts standard | column, bar, line, pie, area, scatter, bubble, radar, stock, combo, axes, series, labels, legend, title, layout, trendlines, error bars | chart helper range parsing, property setters, SVG renderer extraction for standard charts | add/set/get/raw/remove standard chart types and series/axis mutations | all standard chart example scripts plus direct chart CLI |
| Charts extended | waterfall, histogram, boxWhisker and other implemented chartEx types | chartEx resource/remap helpers | add/get/raw/dump/replay extended charts | extended chart example scripts |
| Drawings: pictures | png/jpg/svg input, alt text, crop, hyperlink, anchors, extraction | image type/source helpers and anchor parsing | add/set/get/save picture, save/reopen, raw drawing check | CLI picture add/get --save via shapes/example where present |
| Drawings: shapes | preset geometry, text, font, fill, line, margins, align, valign, gradient | shape property parsing where helperized | add/set/get/remove shapes and raw drawing check | CLI shapes example |
| OLE | embedded package, icon behavior, progId, anchor, binary extraction | OLE helper and invalid prop validation | add/set/get/save OLE, raw rel readback | CLI OLE add/get --save |
| Sparklines | line, column, win/loss, colors, markers, axis, data/location ranges | sparkline range/type validation | add/set/get/remove sparkline groups after reopen | CLI sparklines example |
| Pivot tables | pivot add/get/set/remove, cache, field layout, values, filters, refresh flags | pivot parser/cache/readback/set helpers | pivot lifecycle, cache orphan cleanup, chart/pivot interaction | CLI pivot-tables example |
| Slicers | add/get/set/remove slicers bound to pivot/table where implemented | slicer helper validation where helperized | slicer lifecycle and cache relationship readback | CLI slicers example |
| Query/selectors | CSS-like selectors, Excel cell selector aliases, row predicates, text/regex find, warnings | `AttributeFilter`, selector aliasing, regex prefix, diagnostics | query cells/rows/tables/charts/drawings with selectors | CLI query/get with JSON and text output |
| Raw XML | semantic aliases, zip URI paths, row/col filters, raw-set actions | `RawXmlHelper` actions, namespace/schema-order behavior | raw and raw-set workbook/sheet/chart/drawing parts, validate after mutation | CLI raw/raw-set |
| Dump/batch | dump emitted commands, replay into blank workbook, raw preservation where supported | `ExcelBatchEmitter` element emitters | dump -> batch replay preserves workbook semantics | CLI dump/batch replay |
| Import | CSV/TSV and implemented sheet import paths | parser/options helpers if exposed | import data into workbook and query/readback | CLI import |
| Render HTML | workbook/sheet HTML, styles, formulas, CF visuals, charts, drawings, freeze, overflow | HTML helper, screenshot grid, color/number/CF/sparkline/chart renderer helpers | `ViewAsHtml` contains expected sheet/cell/render structures | CLI `view html --out` |
| Render screenshot | backend discovery and graceful `no_screenshot_backend` behavior | `HtmlScreenshot` dimension/grid/backend helpers | handler-level screenshot path if exposed, otherwise N/A | CLI `view screenshot`, accepting environment-missing backend only |
| Worksheet bloat filter | huge empty cell filtering on open/save | `WorksheetBloatFilter` stream filtering and thresholds | open workbook with bloat cells and verify warning/count/no data loss | CLI validate/open on generated bloat workbook |
| Error handling | bad sheet/range/type/prop, unknown part, invalid formula args, duplicate objects, missing files | validation helpers and parser errors | handler calls throw precise errors and leave workbook valid | CLI non-zero exit and diagnostic text/json |

## CLI Command Matrix

Every Excel command below needs at least one real-process E2E test:

- `create`
- `open`
- `save`
- `close`
- `add`
- `get`
- `query`
- `set`
- `remove`
- `move`
- `swap`
- `batch`
- `dump`
- `raw`
- `raw-set`
- `validate`
- `view html`
- `view screenshot`
- `import`

Commands that are format-generic still count here if Excel has Excel-specific behavior or path syntax.

## Example Matrix

Every deterministic local Excel script must be exercised through E2E by copying it to a temp directory and running it against the current CLI executable:

- `examples/excel/cell-formatting.sh`
- `examples/excel/charts.sh`
- `examples/excel/charts/charts-advanced.sh`
- `examples/excel/charts/charts-area.sh`
- `examples/excel/charts/charts-bar.sh`
- `examples/excel/charts/charts-basic.sh`
- `examples/excel/charts/charts-boxwhisker.sh`
- `examples/excel/charts/charts-bubble.sh`
- `examples/excel/charts/charts-column.sh`
- `examples/excel/charts/charts-combo.sh`
- `examples/excel/charts/charts-extended.sh`
- `examples/excel/charts/charts-histogram.sh`
- `examples/excel/charts/charts-line.sh`
- `examples/excel/charts/charts-pie.sh`
- `examples/excel/charts/charts-radar.sh`
- `examples/excel/charts/charts-scatter.sh`
- `examples/excel/charts/charts-stock.sh`
- `examples/excel/charts/charts-waterfall.sh`
- `examples/excel/conditional-formatting.sh`
- `examples/excel/data-validation.sh`
- `examples/excel/pivot-tables.sh`
- `examples/excel/shapes.sh`
- `examples/excel/sheet-settings.sh`
- `examples/excel/slicers.sh`
- `examples/excel/sparklines.sh`
- `examples/excel/workbook-settings.sh`

## Existing Coverage Status

The current tests are useful but not complete enough for the target:

- Existing `ExcelUnitTests` is representative only; it must be split or expanded into feature-family unit files.
- Existing integration tests cover selected composed workbooks; they do not yet prove every feature area above.
- Existing E2E tests run only a subset of example scripts and commands.
- `EXCEL_COVERAGE.md` lists schema/example/command surfaces, but it does not yet require Unit / Integration / E2E coverage per feature.

## Test File Layout

Keep one test project. Add focused files under `tests/OfficeCli.Tests/Excel`:

- `ExcelFormulaUnitTests.cs`
- `ExcelFormattingUnitTests.cs`
- `ExcelSelectorsAndRawUnitTests.cs`
- `ExcelChartUnitTests.cs`
- `ExcelPivotTableUnitTests.cs`
- `ExcelRenderUnitTests.cs`
- `ExcelWorkbookIntegrationTests.cs`
- `ExcelCellsRangesIntegrationTests.cs`
- `ExcelTablesFiltersValidationIntegrationTests.cs`
- `ExcelConditionalFormattingIntegrationTests.cs`
- `ExcelChartsDrawingsIntegrationTests.cs`
- `ExcelPivotSlicerIntegrationTests.cs`
- `ExcelRawDumpImportIntegrationTests.cs`
- `ExcelCliCommandE2ETests.cs`
- `ExcelExamplesE2ETests.cs`
- `ExcelRenderE2ETests.cs`

This layout avoids one giant test file and maps directly to the matrix.

## Verification Plan

For each implementation batch:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~<batch>"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

After the full matrix is complete:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
./build.sh all
```

## Self-Review

- No placeholders remain.
- Scope is current implemented Excel behavior only.
- The matrix does not mark unimplemented future Excel features as required.
- The plan keeps packaging verification at the end because the matrix build is expensive.
- The file layout adds tests only inside the existing test project.
