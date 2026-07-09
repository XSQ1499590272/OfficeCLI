# OfficeCLI Slimming Design

Date: 2026-07-08

## Goal

Reduce maintenance complexity across the OfficeCLI repository while preserving all public behavior and the Word, Excel, and PowerPoint render-look-fix loop.

The primary goal is maintainability, not aggressive repository-size reduction. Size reduction is acceptable only when it follows from removing duplication, stale entry points, unused assets, or obsolete documentation.

## Non-Negotiable Compatibility

Public entry points stay compatible by default:

- CLI commands and arguments
- `officecli view <file> html`
- `officecli view <file> screenshot`
- `officecli watch <file>`
- install scripts and self-install paths
- npm package name, binary shim, and install behavior
- Python and Node SDK package/API surfaces
- root `SKILL.md` and embedded skill entry points
- published example paths unless an item is proven unused and unreferenced

Any proposed removal of a public entry point requires a separate explicit decision before implementation.

## Render Loop Protection

The core loop is:

1. create or modify a `.docx`, `.xlsx`, or `.pptx`
2. render it to HTML or PNG
3. inspect the rendered output
4. apply a correction
5. re-render or refresh the watch view

The protected commands are:

- `officecli view <file> html`
- `officecli view <file> screenshot`
- `officecli watch <file>`

The protected update paths are `add`, `set`, `remove`, and `batch` notifying a running watch session after document changes.

The headless HTML screenshot path must remain available without requiring Microsoft Office. Windows native rendering may remain as an enhancement, but it cannot become the only rendering path.

## Verification Gates

Each slimming step has two gates.

### Local Gate

Restore the historical local test project at `tests/OfficeCli.Tests`, but keep `tests/` ignored and do not commit the restored test files.

The local test project should be restored from the historical xUnit project introduced around commit `020fc7a1`, with dependencies aligned to the current product dependency versions. In particular, `DocumentFormat.OpenXml` should match the main project version, currently `3.4.1`.

Before implementation begins, `dotnet --info` must work locally. If the .NET SDK is not on `PATH`, the local gate is blocked.

Each step runs:

- local ignored test project
- docx smoke: create, modify, `view html`, `view screenshot`, and validate or inspect issues
- xlsx smoke: create, modify, `view html`, `view screenshot`, and validate or inspect issues
- pptx smoke: create, modify, `view html`, `view screenshot`, and validate or inspect issues
- a watch smoke where practical; if browser interaction is not automatable for the step, keep it as an explicit manual smoke

### CI Gate

After each step, push the branch and run the GitHub Actions cross-platform build matrix. The matrix must prove publish, smoke, and install behavior on macOS, Linux, and Windows.

Also run path-specific checks when relevant:

- `sdk-smoke` when `sdk/`, npm, install, or provisioning paths change
- `skill-parity` when `SKILL.md` or `skills/officecli/SKILL.md` changes
- package publish workflow checks when package metadata changes

Do not use local `./build.sh all` as a replacement for the CI matrix. It can prove RID publish locally, but not real OS behavior.

## Phased Cleanup Order

### Phase 0: Restore Local Guardrails

Restore the ignored local test project and confirm it can run. This phase does not slim product code.

### Phase 1: Structure Consistency

Clean stale or misleading repository structure. This includes solution/test references, build/test documentation, and obvious dead entry points. Keep changes low risk.

### Phase 2: Outer Repository Maintenance Surface

Reduce duplicated or obsolete documentation, examples, and assets. Do not remove published examples or public paths unless they are proven unused and unreferenced.

### Phase 3: SDK, npm, and Install Surface

Remove duplicated or stale logic around provisioning and packaging only when behavior stays compatible. Package names, binary names, asset names, and install locations must remain stable.

### Phase 4: Skills, Schemas, and Examples

Reduce duplicated guidance and obsolete generated artifacts. Preserve embedded resources and root/skill parity. Keep `SKILL.md` and `skills/officecli/SKILL.md` byte-identical when either changes.

### Phase 5: Core Code

Slim `src/officecli` last. Start with dead paths and obvious duplicated helpers. Treat `view`, `screenshot`, `watch`, `RenderViaRegistry`, and the Word/Excel/PowerPoint HTML preview code as high-risk areas and change them only in small, independently verified steps.

## Step Rules

Each step must be independently reviewable, reversible, and gated.

1. Make the smallest change that advances the current phase.
2. Run the local gate.
3. Push and run the CI gate.
4. Continue only after both gates pass.
5. If a gate fails, stop slimming and fix only the regression or stale test problem caused by the current step.

Do not combine unrelated cleanup areas in one step.

## Explicit Non-Goals

- Do not redesign the renderer architecture as part of this cleanup.
- Do not replace the render-loop commands.
- Do not add new frameworks or dependencies for cleanup.
- Do not commit the ignored local `tests/` directory.
- Do not break public CLI, SDK, install, skill, or example entry points for cosmetic structure.
- Do not use CI failures as permission for broad refactoring.

## Success Criteria

The slimming work is successful when:

- each phase produces small, reviewable diffs
- every step passes the local ignored test gate
- every step passes the GitHub Actions cross-platform CI gate
- Word, Excel, and PowerPoint still support render-look-fix through HTML, screenshot, and watch
- public entry points remain compatible
- the final repository has fewer stale paths, duplicated maintenance surfaces, or dead branches than it started with
