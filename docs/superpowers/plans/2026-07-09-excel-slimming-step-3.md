# Excel Slimming Step 3 Plan

## Goal

Remove duplicated Excel drawing anchor constants without changing behavior.

## Scope

- Reuse the existing class-level `EmuPerColApprox` and `EmuPerRowApprox` constants in `ParseAnchorOrigin`.
- Reuse the same constants in `ParseAnchorDimension`.
- Keep parsing behavior, validation, and error messages unchanged.

## Out Of Scope

- No public CLI command, argument, schema key, compatibility alias, renderer entrypoint, watch path, or packaging script changes.
- No Word or PowerPoint source changes.
- No directory moves in this step.

## Verification

1. `dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --nologo`
