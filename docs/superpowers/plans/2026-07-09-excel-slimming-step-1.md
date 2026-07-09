# Excel Slimming Step 1 Plan

## Goal

Remove a first small set of Excel-only dead code while preserving the Word/Excel/PowerPoint render-look-fix loop and cross-platform packaging behavior.

## Scope

- Remove `ExcelHandler.GetGlobalChartPart`, which has no callers.
- Remove `ExcelHandler.RenderSheetCharts`, which is an unused wrapper around `CollectSheetCharts`; the active HTML preview path calls `CollectSheetCharts` directly.
- Remove `ExcelHandler.IsCanonicalNumericText`, which has no callers. Keep `CanonicalNumericLiteral` because `NormalizeNumericCellText` still uses it.

## Out Of Scope

- No public CLI command, argument, schema key, compatibility alias, renderer, watch path, or packaging script changes.
- No Word or PowerPoint source changes.
- No directory moves in this step.

## Verification

1. `dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo`
2. `./build.sh all`
