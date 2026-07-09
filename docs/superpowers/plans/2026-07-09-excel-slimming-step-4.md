# Excel Slimming Step 4 Plan

## Goal

Remove duplicated Excel drawing anchor grid-limit constants without changing behavior.

## Scope

- Replace the duplicated local `MaxCellIndex = 16384` constants in drawing anchor parsing with one class-level constant.
- Keep the existing validation boundary and error messages equivalent.

## Out Of Scope

- No public CLI command, argument, schema key, compatibility alias, renderer entrypoint, watch path, or packaging script changes.
- No Word or PowerPoint source changes.
- No directory moves in this step.

## Verification

1. `dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo`
