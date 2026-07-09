# Excel Unit Integration E2E Tests Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Complete Excel test coverage as three explicit layers: unit, integration, and e2e.

**Architecture:** Keep the existing xUnit project and `tests/OfficeCli.Tests/Excel/ExcelTestBase.cs`. Add the smallest missing test files under `tests/OfficeCli.Tests/Excel/`; only add test-only internal visibility if unit tests need internal helpers. Keep CLI/render checks in e2e tests so handler tests stay fast.

**Tech Stack:** .NET 10, xUnit, FluentAssertions, existing `OfficeCli.Handlers.ExcelHandler`, existing CLI publish/build scripts.

## Global Constraints

- Do not add a second test project unless `tests/OfficeCli.Tests` cannot express the layer.
- Do not add new NuGet packages.
- Production-code edits are limited to test visibility (`InternalsVisibleTo`) unless a test exposes a real bug and the user approves the fix.
- Every test-authoring task ends with `dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo`.
- Run `./build.sh all` once after all Excel unit, integration, and e2e work is complete, before claiming the full task is done.
- Screenshot e2e must accept `no_screenshot_backend` as an environment limitation, but `view html` must always pass.

---

## Current Code Facts

- `tests/OfficeCli.Tests/OfficeCli.Tests.csproj` already targets `net10.0`, references `src/officecli/officecli.csproj`, and uses xUnit + FluentAssertions.
- Current Excel tests live in `tests/OfficeCli.Tests/Excel/` and contain 11 `[Fact]` tests.
- Existing Excel tests are mostly direct `ExcelHandler` round-trip tests plus one CLI/render smoke test.
- `src/officecli/Handlers/ExcelHandler.cs` exposes public `Raw`, `Validate`, `ViewAsText`, `ViewAsHtml`, `Import`, and mutation methods through `IDocumentHandler`.
- Many good unit targets are `internal`, including `FormulaEvaluator`, `ExcelDataFormatter`, `HtmlScreenshot`, `AttributeFilter`, and Excel helper methods.
- There is no current `InternalsVisibleTo` entry for `OfficeCli.Tests`.

## File Structure

- Create: `src/officecli/Properties/AssemblyInfo.cs`
  - Test-only access to internal helpers.
- Create: `tests/OfficeCli.Tests/Excel/ExcelUnitTests.cs`
  - Pure helper/formula/selector/format tests.
- Create: `tests/OfficeCli.Tests/Excel/ExcelIntegrationTests.cs`
  - Public `ExcelHandler` end-to-end handler tests that compose multiple features in one workbook.
- Create: `tests/OfficeCli.Tests/Excel/ExcelE2ETests.cs`
  - Real CLI process tests for create/open/batch/dump/raw/view/validate and examples.
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`
  - Add a Layer Surface section mapping Unit / Integration / E2E.

## Task 1: Enable Unit Testing Of Internal Excel Helpers

**Files:**
- Create: `src/officecli/Properties/AssemblyInfo.cs`

**Steps:**

- [ ] Add this file:

```csharp
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("OfficeCli.Tests")]
```

- [ ] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: tests pass.

## Task 2: Add Excel Unit Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelUnitTests.cs`

**Minimum tests:**

- `FormulaEvaluator_EvaluatesRepresentativeFamilies`
  - Build a tiny in-memory worksheet and assert arithmetic, lookup, date/text, dynamic array, financial, statistical, securities where currently supported.
- `FormulaEvaluator_ReturnsExcelErrorsForInvalidMath`
  - Assert `SQRT(-1)`, `LOG(0)`, bad refs, divide by zero report Excel-style errors.
- `ExcelDataFormatter_FormatsDatesPercentCurrencyAndText`
  - Assert built-in and custom number formats.
- `AttributeFilter_ParsesAndAppliesExcelCellSelectors`
  - Assert equality, contains, numeric comparisons, regex text filter, `and/or`.
- `ExcelHelperValidation_RejectsBadSheetNamesRangesAndCellValues`
  - Assert bad sheet names, bad sqref, invalid sparkline range, too-long cell text.
- `HtmlScreenshot_AutoGridColumns_IsStable`
  - Assert column choice for 1, 2, 4, 9 items without requiring a browser.

**Steps:**

- [ ] Add the tests above, using internal helpers directly where possible.
- [ ] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelUnitTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: unit filter passes and full suite passes.

## Task 3: Add Excel Integration Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelIntegrationTests.cs`

**Minimum tests:**

- `WorkbookAuthoring_ComposesTablesChartsCfValidationAndFormulas`
  - One workbook with table, formulas, validation, conditional formatting, chart, named range; save, reopen, get/query/validate.
- `WorkbookMutation_RewritesReferencesAcrossSheetRowColumnMoves`
  - Insert/move/remove rows and columns, then assert formulas, named ranges, chart ranges, print areas still read back correctly.
- `WorkbookDrawingRoundTrip_PreservesPicturesShapesOleAndSparklines`
  - Add drawing objects, save/reopen, assert raw drawing part and semantic query both survive.
- `WorkbookDumpBatchReplay_RecreatesSemantics`
  - Use `ExcelBatchEmitter.EmitExcel` or CLI dump output, replay into a blank workbook, compare key `Get` nodes.

**Steps:**

- [ ] Add the tests above using `ExcelTestBase`.
- [ ] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelIntegrationTests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: integration filter passes and full suite passes.

## Task 4: Add Excel E2E Tests

**Files:**
- Create: `tests/OfficeCli.Tests/Excel/ExcelE2ETests.cs`

**Minimum tests:**

- `CliExcel_CreateOpenBatchSaveCloseValidate`
  - Spawn the built `officecli.dll`; run create, open, batch, save, close, validate.
- `CliExcel_DumpBatchRawReplay`
  - CLI dump a populated workbook, batch replay into a blank workbook, raw read workbook and sheet XML.
- `CliExcel_ViewHtmlAlwaysRendersWorkbook`
  - `view html --out` creates HTML containing sheet/cell content and no `###` for controlled widths.
- `CliExcel_ViewScreenshotRendersOrReportsMissingBackend`
  - Accept PNG output or `no_screenshot_backend`.
- `CliExcel_ExamplesSmoke`
  - Run existing `examples/excel/*.sh` and `examples/excel/charts/*.sh` scripts that are deterministic and local-only.

**Steps:**

- [ ] Move the existing CLI helper from `ExcelDumpBatchRenderTests` into `ExcelTestBase` if reuse keeps code shorter.
- [ ] Add the tests above.
- [ ] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo --filter "FullyQualifiedName~ExcelE2ETests"
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
```

Expected: e2e filter passes except browser-dependent screenshot may assert `no_screenshot_backend`; full suite passes.

## Task 5: Update Coverage Manifest

**Files:**
- Modify: `tests/OfficeCli.Tests/Excel/EXCEL_COVERAGE.md`

**Steps:**

- [ ] Add this section:

```markdown
## Layer Surface

- [x] Unit -> ExcelUnitTests
- [x] Integration -> ExcelIntegrationTests
- [x] E2E -> ExcelE2ETests
```

- [ ] Keep the existing Schema / Example / Command surfaces checked only where tests still exist.
- [ ] Run:

```bash
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo
./build.sh all
```

Expected: full suite passes and all 8 release binaries build.

## Code-Fact Review

- The plan matches the current test project: no new test project or package is needed.
- `InternalsVisibleTo` is required for real unit tests because the code facts show no existing friend assembly and many unit targets are `internal`.
- Current Excel tests already cover schema/command breadth, but they do not explicitly separate Unit / Integration / E2E layers.
- The plan keeps render-look-fix coverage in e2e via `view html` and `view screenshot`.
- The plan keeps cross-platform packaging verification on the existing `./build.sh all` path.
- No contradiction found against the current code layout.
