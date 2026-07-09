# Excel Slimming Step 2 Plan

## Goal

Continue a small, Excel-only code cleanup without changing Word, Excel, or PowerPoint behavior.

## Scope

- Inline the single-use `ChartColOffsetPt` helper into `CollectSheetCharts`.
- Keep the existing partial-column EMU offset calculation unchanged.

## Out Of Scope

- No public CLI command, argument, schema key, compatibility alias, renderer entrypoint, watch path, or packaging script changes.
- No Word or PowerPoint source changes.
- No directory moves in this step.

## Verification

1. `dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo`
