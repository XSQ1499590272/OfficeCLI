# Excel Test Coverage

Every checked item must have at least one tracked test before Excel code slimming starts.
Coverage is tracked across three layers: Unit (pure logic/error branches), Integration (handler API with save/reopen readback), E2E (real CLI process).

Total: **347 tests, 0 failures** (as of 2026-07-09).

## Schema Surface

| Surface | Unit | Integration | E2E |
| --- | --- | --- | --- |
| aboveaverage | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| autofilter | N/A (direct Open XML orchestration) | ExcelTablesFiltersValidationIntegrationTests | ExcelCliCommandE2ETests |
| cell | N/A (direct Open XML orchestration) | ExcelCellsRangesIntegrationTests, ExcelWorkbookIntegrationTests | ExcelCliCommandE2ETests |
| cellis | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| cfextended | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| chart | ExcelChartUnitTests | ExcelChartsDrawingsIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests, ExcelExamplesE2ETests |
| chart-axis | ExcelChartUnitTests | ExcelChartsDrawingsIntegrationTests | ExcelCliCommandE2ETests |
| chart-series | ExcelChartUnitTests | ExcelChartsDrawingsIntegrationTests | ExcelCliCommandE2ETests |
| colbreak | N/A (direct Open XML orchestration) | ExcelWorkbookIntegrationTests | ExcelCliCommandE2ETests |
| colorscale | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| column | N/A (direct Open XML orchestration) | ExcelCellsRangesIntegrationTests | ExcelCliCommandE2ETests |
| comment | N/A (direct Open XML orchestration) | ExcelCellsRangesIntegrationTests | ExcelCliCommandE2ETests |
| conditionalformatting | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| containstext | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| databar | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| dateoccurring | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| detectedtable | N/A (direct Open XML orchestration) | ExcelTablesFiltersValidationIntegrationTests | ExcelCliCommandE2ETests |
| duplicatevalues | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| formulacf | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| hyperlink | N/A (direct Open XML orchestration) | ExcelCellsRangesIntegrationTests | ExcelCliCommandE2ETests |
| iconset | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| namedrange | N/A (direct Open XML orchestration) | ExcelTablesFiltersValidationIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests |
| ole | N/A (direct Open XML orchestration) | ExcelChartsDrawingsIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests |
| pagebreak | N/A (direct Open XML orchestration) | ExcelWorkbookIntegrationTests | ExcelCliCommandE2ETests |
| picture | N/A (direct Open XML orchestration) | ExcelChartsDrawingsIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests |
| pivottable | ExcelPivotTableUnitTests | ExcelPivotSlicerIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests, ExcelExamplesE2ETests |
| range | N/A (direct Open XML orchestration) | ExcelCellsRangesIntegrationTests | ExcelCliCommandE2ETests |
| raw | ExcelSelectorsAndRawUnitTests | ExcelRawDumpImportIntegrationTests | ExcelCliCommandE2ETests, ExcelE2ETests |
| row | N/A (direct Open XML orchestration) | ExcelCellsRangesIntegrationTests | ExcelCliCommandE2ETests |
| rowbreak | N/A (direct Open XML orchestration) | ExcelWorkbookIntegrationTests | ExcelCliCommandE2ETests |
| run | N/A (direct Open XML orchestration) | ExcelCellsRangesIntegrationTests | ExcelCliCommandE2ETests |
| shape | N/A (direct Open XML orchestration) | ExcelChartsDrawingsIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests, ExcelExamplesE2ETests |
| sheet | N/A (direct Open XML orchestration) | ExcelWorkbookIntegrationTests | ExcelCliCommandE2ETests, ExcelExamplesE2ETests |
| slicer | N/A (direct Open XML orchestration) | ExcelPivotSlicerIntegrationTests, ExcelIntegrationTests | ExcelExamplesE2ETests |
| sort | N/A (direct Open XML orchestration) | ExcelTablesFiltersValidationIntegrationTests | ExcelCliCommandE2ETests |
| sparkline | N/A (direct Open XML orchestration) | ExcelChartsDrawingsIntegrationTests, ExcelIntegrationTests | ExcelExamplesE2ETests |
| table | N/A (direct Open XML orchestration) | ExcelTablesFiltersValidationIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests |
| topn | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| uniquevalues | N/A (direct Open XML orchestration) | ExcelConditionalFormattingIntegrationTests | ExcelCliCommandE2ETests |
| validation | N/A (direct Open XML orchestration) | ExcelTablesFiltersValidationIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests, ExcelExamplesE2ETests |
| workbook | N/A (direct Open XML orchestration) | ExcelWorkbookIntegrationTests | ExcelCliCommandE2ETests, ExcelExamplesE2ETests |

## Example Surface

Example-surface E2E counts only when a current test asserts example-specific artifacts, not when it only runs the script and validates the workbook.

| Surface | Unit | Integration | E2E |
| --- | --- | --- | --- |
| cell-formatting | N/A | ExcelCellsRangesIntegrationTests | ExcelExamplesE2ETests |
| charts/* (all 16 chart sub-scripts) | ExcelChartUnitTests | ExcelChartsDrawingsIntegrationTests | ExcelExamplesE2ETests |
| conditional-formatting | N/A | ExcelConditionalFormattingIntegrationTests | ExcelExamplesE2ETests |
| data-validation | N/A | ExcelTablesFiltersValidationIntegrationTests | ExcelExamplesE2ETests |
| pivot-tables | ExcelPivotTableUnitTests | ExcelPivotSlicerIntegrationTests | ExcelExamplesE2ETests |
| shapes | N/A | ExcelChartsDrawingsIntegrationTests | ExcelExamplesE2ETests |
| sheet-settings | N/A | ExcelWorkbookIntegrationTests | ExcelExamplesE2ETests |
| slicers | N/A | ExcelPivotSlicerIntegrationTests | ExcelExamplesE2ETests |
| sparklines | N/A | ExcelChartsDrawingsIntegrationTests | ExcelExamplesE2ETests |
| workbook-settings | N/A | ExcelWorkbookIntegrationTests | ExcelExamplesE2ETests |

## Command Surface

| Surface | Unit | Integration | E2E |
| --- | --- | --- | --- |
| create | N/A (filesystem operation) | ExcelWorkbookIntegrationTests | ExcelCliCommandE2ETests, ExcelE2ETests |
| add | N/A (router dispatch) | All Integration test classes | ExcelCliCommandE2ETests, ExcelE2ETests |
| get | N/A (router dispatch) | All Integration test classes | ExcelCliCommandE2ETests, ExcelE2ETests |
| query | ExcelSelectorsAndRawUnitTests | All Integration test classes | ExcelCliCommandE2ETests, ExcelE2ETests |
| set | N/A (router dispatch) | All Integration test classes | ExcelCliCommandE2ETests, ExcelE2ETests |
| remove | N/A (router dispatch) | ExcelCellsRangesIntegrationTests, ExcelWorkbookIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests |
| move | N/A (router dispatch) | ExcelCellsRangesIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests |
| swap | N/A (router dispatch) | N/A | ExcelCliCommandE2ETests |
| batch | N/A (router dispatch) | ExcelRawDumpImportIntegrationTests | ExcelCliCommandE2ETests, ExcelE2ETests |
| dump | N/A (router dispatch) | ExcelRawDumpImportIntegrationTests | ExcelCliCommandE2ETests, ExcelE2ETests |
| raw | ExcelSelectorsAndRawUnitTests | ExcelRawDumpImportIntegrationTests | ExcelCliCommandE2ETests, ExcelE2ETests |
| raw-set | ExcelSelectorsAndRawUnitTests | ExcelRawDumpImportIntegrationTests | ExcelCliCommandE2ETests |
| validate | N/A (router dispatch) | All Integration test classes | ExcelCliCommandE2ETests, ExcelE2ETests |
| view html | ExcelRenderUnitTests | ExcelIntegrationTests | ExcelRenderE2ETests |
| view screenshot | ExcelRenderUnitTests | N/A (backend-dependent) | ExcelRenderE2ETests |
| import | N/A (router dispatch) | ExcelRawDumpImportIntegrationTests, ExcelIntegrationTests | ExcelCliCommandE2ETests |

## Layer Surface

| Test Class | Unit | Integration | E2E |
| --- | --- | --- | --- |
| ExcelUnitTests | x |  |  |
| ExcelFormulaUnitTests | x |  |  |
| ExcelFormattingUnitTests | x |  |  |
| ExcelSelectorsAndRawUnitTests | x |  |  |
| ExcelChartUnitTests | x |  |  |
| ExcelPivotTableUnitTests | x |  |  |
| ExcelRenderUnitTests | x |  |  |
| ExcelCellsAndFormulasTests (legacy) |  | x |  |
| ExcelChartsAndDrawingsTests (legacy) |  | x |  |
| ExcelConditionalFormattingTests (legacy) |  | x |  |
| ExcelIntegrationTests (legacy) |  | x |  |
| ExcelPivotAndSlicerTests (legacy) |  | x |  |
| ExcelTablesFiltersValidationTests (legacy) |  | x |  |
| ExcelWorkbookIntegrationTests |  | x |  |
| ExcelCellsRangesIntegrationTests |  | x |  |
| ExcelTablesFiltersValidationIntegrationTests |  | x |  |
| ExcelConditionalFormattingIntegrationTests |  | x |  |
| ExcelChartsDrawingsIntegrationTests |  | x |  |
| ExcelPivotSlicerIntegrationTests |  | x |  |
| ExcelRawDumpImportIntegrationTests |  | x |  |
| ExcelDumpBatchRenderTests (legacy) |  |  | x |
| ExcelE2ETests (legacy) |  |  | x |
| ExcelCliCommandE2ETests |  |  | x |
| ExcelExamplesE2ETests |  |  | x |
| ExcelRenderE2ETests |  |  | x |

## Unit Test Details

### ExcelFormulaUnitTests (Task 2)
- FormulaEvaluator dispatch families: arithmetic, lookup, text, date/time, dynamic array/spill, financial, statistics, error propagation
- ModernFunctionQualifier: _xlfn qualification, dynamic array detection, sheet name quoting
- FormulaRefShifter: row/column/copy/sheet-rename rewriting, absolute/relative refs, structured text preservation
- FormulaCacheHelpers: allowlist classification, computed/cache agreement with numeric tolerance

### ExcelFormattingUnitTests (Task 3)
- ExcelDataFormatter: general, integer, decimal, percent, scientific, fraction, currency/accounting, date, time, text, custom section formats
- Number-format section selection (positive/negative/zero/text)

### ExcelSelectorsAndRawUnitTests (Task 3)
- AttributeFilter: equality, inequality, contains, numeric comparison, exists, regex find, and/or, unknown-key diagnostics
- RawXmlHelper: append, prepend, insertbefore, insertafter, replace, remove, setattr actions
- Namespace preservation, schema-order insertion, zip URI path detection
- WorksheetBloatFilter: empty-cell filtering while preserving value/formula/style cells
- Excel cell alias normalization

### ExcelChartUnitTests (Task 4)
- ChartHelper: range parsing, series refs, chart type aliases, axis normalization, title/legend/label setters, trendline/error-bar options, secondary axis mapping
- ChartExBuilder: extended chart resource loading, range remapping
- ChartSvgRenderer: standard chart data extraction, per-point colors, deleted labels, marker settings, trendlines, error bars, reference lines

### ExcelPivotTableUnitTests (Task 4)
- PivotTableHelper: cache parsing, field/value/filter parsing, settable properties, refresh flags, orphan cache pruning

### ExcelRenderUnitTests (Task 3)
- HtmlScreenshot: grid sizing, backend detection handling, color normalization, number-format section selection

## Verification

```bash
# Full suite (targeted filters for each batch):
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelFormulaUnitTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelFormattingUnitTests|FullyQualifiedName~ExcelSelectorsAndRawUnitTests|FullyQualifiedName~ExcelRenderUnitTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelChartUnitTests|FullyQualifiedName~ExcelPivotTableUnitTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelWorkbookIntegrationTests|FullyQualifiedName~ExcelCellsRangesIntegrationTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelTablesFiltersValidationIntegrationTests|FullyQualifiedName~ExcelConditionalFormattingIntegrationTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelChartsDrawingsIntegrationTests|FullyQualifiedName~ExcelPivotSlicerIntegrationTests|FullyQualifiedName~ExcelRawDumpImportIntegrationTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelCliCommandE2ETests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelExamplesE2ETests|FullyQualifiedName~ExcelRenderE2ETests"

# Final: full suite + release build
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
./build.sh all
```
