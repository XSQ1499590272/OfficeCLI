# PPT Unit And E2E Test Implementation Plan

> **Status: COMPLETE** — All 10 tasks implemented, reviewed, and verified.
> **Final results:** 1,203 tests passing, 0 failures, `./build.sh all` passes (8 release binaries).
> **Branch:** `dev` | **Commits:** `5e3b8f3f..6513eb11` (30 commits)

**Goal:** Complete Unit/direct-handler and E2E tests for every currently implemented or exposed PowerPoint feature in OfficeCLI.

**Architecture:** Keep the existing `OfficeCli.Tests` xUnit project and add focused PPT tests under `tests/OfficeCli.Tests/Pptx`. Unit tests cover helper logic and direct `PowerPointHandler` behavior without a process boundary; E2E tests use the real CLI only where command routing, resident behavior, examples, rendering, dump/replay, or user-facing errors are the behavior under test.

**Tech Stack:** .NET 10, xUnit, FluentAssertions, DocumentFormat.OpenXml, existing `OfficeCli.Handlers.PowerPointHandler`, existing `PptxBatchEmitter`, existing `CommandBuilder`, existing `BlankDocCreator`, existing `./build.sh all`.

## Global Constraints

- Scope is current implemented or exposed PPT/PPTX behavior only; do not add new PowerPoint product features.
- Do not add a second test project.
- Do not add new NuGet packages.
- Do not add automatic commit steps; the user owns final submission.
- Production-code edits are allowed only for test visibility or for a real bug exposed by tests after user approval.
- Each implementation batch must run its filtered tests and then the full `OfficeCli.Tests` suite.
- Run `./build.sh all` once after the full PPT matrix is complete.
- Screenshot E2E may accept `no_screenshot_backend`; `view html` and `view svg` must produce inspectable output.
- Native screenshot assertions must be conditional: `--render native` is Windows + Microsoft PowerPoint only.

---

## File Structure

- Created: `tests/OfficeCli.Tests/Pptx/PptTestBase.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PPTX_COVERAGE.md` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptPathSelectorRawUnitTests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptTextShapeTableUnitTests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptMediaChartEffectUnitTests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptRenderDumpUnitTests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptPresentationSlideHandlerTests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptContentHandlerTests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptAdvancedHandlerTests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptCliCommandE2ETests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptExamplesE2ETests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptRenderE2ETests.cs` ✅
- Created: `tests/OfficeCli.Tests/Pptx/PptDumpReplayE2ETests.cs` ✅

## Task 1: Coverage Manifest And Shared Test Helpers ✅ COMPLETE

**Commit:** `5a41128c`
**Review:** ✅ Clean

- [x] Add `PptTestBase` by copying the working shape of `ExcelTestBase` and changing only the document type, temp prefixes, and handler type.
- [x] `CreatePresentation()` calls `BlankDocCreator.Create(path)` for a `.pptx`.
- [x] `RunCli` sets `OFFICECLI_NO_AUTO_RESIDENT=1`, matching the Excel tests.
- [x] Add `PPTX_COVERAGE.md` with rows for all 34 PPT feature areas.
- [x] All 15 helper methods implemented (CreatePresentation, OpenEditable, OpenReadOnly, TrackTempDirectory, NewTempPath, ReadZipEntry, RunCli, RunCliOk, CopyExampleScriptToTemp, RunShellScriptOk, CreateTinyPng, CreateTinySvg, CreateOlePayload, CreateTinyGlb, CreateTinyVideo).
- [x] Full suite passes (491 tests at time of completion).

## Task 2: Path, Selector, Raw, And Schema Unit Tests ✅ COMPLETE

**Commits:** `52b9cd55`, `bbc29f85` (fix)
**Review:** ✅ Clean after 1 fix round
**Tests:** 56

- [x] Selector tests for type selectors, attribute equality/inequality/contains/exists/numeric comparison, nested group traversal, full-path `@name=` connector targets, and unknown-key diagnostics.
- [x] Path tests for `/slide[N]`, `/slide[N]/shape[N]`, `/slide[N]/group[N]/shape[N]`, `/slide[N]/table[N]/row[N]/cell[N]`, `/slide[N]/notes`, `/slideMaster[N]`, `/slideLayout[N]`.
- [x] Raw XML tests for `raw` and `raw-set` actions: `append`, `prepend`, `insertbefore`, `insertafter`, `replace`, `remove`, `setattr`.
- [x] Raw path tests for `/presentation`, `/slide[N]`, `/slideMaster[N]`, `/slideLayout[N]`, `/theme`, `/noteSlide[N]`, `/chart[N]`.
- [x] Help/schema tests that every `schemas/help/pptx/*.json` element can be rendered by `officecli help pptx <element> --json`.
- [x] Negative tests for invalid slide index, missing raw part, invalid XPath, invalid property, and unsupported element/type names.

## Task 3: Text, Shape, Connector, Group, Placeholder, And Table Unit Tests ✅ COMPLETE

**Commits:** `2d3c4b7a`, `fcf91f13` (fix r1), `e5709c6c` (fix r2)
**Review:** ✅ Clean after 2 fix rounds
**Tests:** 120

- [x] Text tests for textbox, paragraph, run, linebreak, equation, RTL direction, language, font properties, size, color, fill, bold, italic, underline, underline.color, strike, highlight, baseline, spacing, caps, and hyperlinks.
- [x] Shape tests for preset geometry, arrow alias, custom geometry paths, x/y/width/height, rotation, name, alt text, fill, gradient, line, opacity, shadow, glow, blur, z-order, align, distribute, and effective property readback.
- [x] Connector tests for `from`/`to` full paths, `fromSide`, `toSide`, `fromIdx`, `toIdx`, edge-to-edge defaults, connector text labels, and invalid bare `@name=` rejection.
- [x] Group tests for add/query/get/remove nested content, group move/copy, group link/tooltip, and `ungroup=true`.
- [x] Placeholder tests for add/set/remove by `phType`, slide layout placeholder inheritance, and query/readback.
- [x] Table tests for add/get/query/set/remove, row/column/cell addressing, merged cells, fills/backgrounds, borders, built-in styles, row/column move/copyFrom, virtual `/col[C]`, RTL cell text, margins, horizontal/vertical alignment.

## Task 4: Presentation, Slide, Theme, Notes, And Comments Handler Tests ✅ COMPLETE

**Commits:** `efa5a90f`, `dd29ed94` (fix)
**Review:** ✅ Clean after 1 fix round
**Tests:** 95

- [x] Presentation lifecycle tests for create/open/save/reopen/validate, slide size, document properties, presentation settings, default theme, read-only mutation-throws.
- [x] Slide lifecycle tests for add, duplicate/copy, remove, move, swap, hidden, layout, background, notes, header/footer/date/slidenum toggles, and slide-level hyperlinks.
- [x] slideMaster and slideLayout tests for get/query/add/set/remove paths.
- [x] Theme tests for readback, settable theme properties (all 12 color slots + fonts), and HTML/SVG-visible theme color resolution.
- [x] Notes tests for text, paragraph/run formatting, RTL direction, language, and save/reopen.
- [x] Legacy comment and modern comment tests for add/get/query/set/remove, author initials, timestamps, threaded modern comments, RTL text, and dump readback.

## Task 5: Media, Charts, Animations, Transitions, Zoom, OLE, Model3D, And Diagram Tests ✅ COMPLETE

**Commit:** `27a26b3f`
**Review:** ✅ Approved (non-blocking minor issues noted)
**Tests:** 213

- [x] Picture tests for PNG/SVG input, crop/cropleft/croptop/cropright/cropbottom, fill modes, brightness, contrast, shadow, glow, rotation, link, tooltip, alt text.
- [x] Video/audio tests for add/get/set/remove, loop, autoStart, poster/preview where exposed, and relationship/content-type readback.
- [x] OLE tests for add/get/save/reopen, preview image, payload extraction, relationship and content-type readback.
- [x] Model3D tests for `.glb`, rotation, camera/zoom properties where exposed, save/reopen, and dump readback.
- [x] Chart tests for column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, combo, 3D, waterfall, series add/remove, axis properties, gridlines, labels, legend, title, trendlines, error bars, anchor shorthand, and chart XML readback.
- [x] Animation tests for entrance/default path, 15 emphasis presets, 16 exit presets, multi-effect chains, motion-path presets, repeat, restart, autoReverse, chartBuild byCategory/bySeries.
- [x] Transition tests for basic transitions, 12 p15 presets, p14/morph, duration, advanceOnClick, advanceAfter, sound where currently supported.
- [x] Zoom tests for slide zoom add/get/set/remove and target relationship readback.
- [x] Diagram tests for mermaid flowchart/native-shape output and rendered-image fallback.

## Task 6: Render, SVG, Screenshot, Dump, And Batch Unit Tests ✅ COMPLETE

**Commits:** `f7b3fe98`, `2042f580` (fix)
**Review:** ✅ Clean after 1 fix round
**Tests:** 72

- [x] HTML preview tests for slides, shapes, grouped shapes, tables, pictures, charts, notes, comments, hyperlinks, text effects, RTL text, theme colors, and hidden slides.
- [x] SVG preview tests for slide geometry, text, shapes, pictures, charts where rendered, and strict page handling.
- [x] Screenshot helper tests for default first-slide behavior, `--page`, page ranges, `--grid auto`, forced grid columns, viewport sizing, and `--render html` fallback.
- [x] Dump emitter tests for full deck, `/slide[N]` subtree, charts, media, notes, comments, animations, transitions, OLE, model3d, SmartArt/add-part/raw-set carrier behavior.
- [x] Batch behavior tests for continue-on-error, stop-on-error, warnings, JSON envelope shape, and invalid command diagnostics.

## Task 7: Complete CLI Command E2E Tests ✅ COMPLETE

**Commits:** `84cedf48`, `5370ac67` (fix)
**Review:** ✅ Clean after 1 fix round
**Tests:** 73

- [x] Real-process E2E tests for `create`, `open`, `save`, `close`, `add`, `get`, `query`, `set`, `remove`, `move`, `swap`, `batch`, `dump`, `raw`, `raw-set`, `add-part`, `validate`, `view text`, `view outline`, `view stats`, `view issues`.
- [x] Text output and JSON output for `get`, `query`, `view`, errors, warnings, and batch results.
- [x] Non-zero CLI errors for missing file, invalid slide/path/type/prop, invalid relationship target, invalid raw part, invalid XPath, unsupported chart type, invalid media source, and duplicate/ambiguous object.
- [x] Every mutating command validates with `validate`.
- [x] Resident flow: `open` then multiple mutations then `save` then `close`; disk readback verified after `save` and after `close`.

## Task 8: Complete Examples E2E Tests ✅ COMPLETE

**Commits:** `3addd421`, `a169c759` (fix)
**Review:** ✅ Clean after 1 fix round
**Tests:** 40

- [x] E2E coverage for deterministic scripts under `examples/ppt/*.sh`, `examples/ppt/charts/*.sh`, `examples/ppt/shapes/*.sh`, `examples/ppt/tables/*.sh`, `examples/ppt/textboxes/*.sh`, `examples/ppt/transitions/*.sh`.
- [x] Quarantined groups with documented blockers: `pictures` (needs Pillow), `ole` (needs Pillow), `video` (needs imageio/ffmpeg), `3d-model` (needs 4.3MB model file), `templates` (needs pre-existing .pptx files).
- [x] Scripts run in temp directories, temp `officecli` wrapper at front of `PATH`.
- [x] Every generated `.pptx` validated, smoke readbacks verify slide count and content.
- [x] All quarantines documented in `PPTX_COVERAGE.md` with exact blockers.

## Task 9: Render And Dump/Replay E2E Tests ✅ COMPLETE

**Commits:** `fc0f2124`, `c14cd09b` (fix)
**Review:** ✅ Approved
**Tests:** 43 (26 render + 17 dump/replay)

- [x] `view html --out` E2E tests for text, shapes, pictures, tables, charts, notes, hyperlinks, RTL text, theme colors, hidden slides, and animation markers (scaffold).
- [x] `view svg --out` E2E tests for single slide and page-range behavior.
- [x] `view screenshot` E2E tests accepting PNG file or `no_screenshot_backend`; PNG signature verified when produced.
- [x] `view screenshot --grid auto` and `--grid 3` tests for multi-slide decks.
- [x] `view screenshot --render html` tests working without native PowerPoint.
- [x] `view screenshot --render native` tests behind Windows + PowerPoint availability probe.
- [x] Full dump/replay E2E: create rich deck → dump → batch replay → validate → compare essential nodes.
- [x] Subtree dump/replay E2E for `/slide[N]`, `/theme`, `/noteSlide[N]`, `/slideMaster[N]`, `/slideLayout[N]`, `/presentation`.

## Task 10: Coverage Manifest Closure And Final Verification ✅ COMPLETE

**Commit:** `9816a8e7`
**Verification:** ✅ Full suite passes, `./build.sh all` passes

- [x] Every `PPTX_COVERAGE.md` matrix row updated with exact test class and method names for Unit/direct-handler and E2E.
- [x] `N/A` only where feature is not implemented or cannot run in CI, with documented reasons.
- [x] No placeholder coverage claims (TBD/TODO search returned zero results).
- [x] Full suite: 1,203 tests passing, 0 failures.
- [x] `./build.sh all` passes: all 8 release binaries build (mac-arm64, mac-x64, linux-x64, linux-arm64, linux-alpine-x64, linux-alpine-arm64, win-x64, win-arm64).

---

## Self-Review

- Spec coverage: every PPT schema element under `schemas/help/pptx/*.json` maps to at least one test. ✅
- CLI coverage: every command that supports `.pptx` is covered in Task 7 or Task 9. ✅
- Example coverage: deterministic `examples/ppt/**/*.sh` scripts are covered in Task 8; 5 groups quarantined with documented reasons. ✅
- Render coverage: HTML, SVG, screenshot, grid screenshot, and native-conditional screenshot paths covered in Task 9. ✅
- Dump/replay coverage: full deck, subtree, raw/add-part, media, chart, notes, comments, animation, transition, OLE, model3d covered in Task 6 and Task 9. ✅
- Scope control: no new packages, no new test project, no automatic commit steps. ✅

## Final Metrics

| Metric | Count |
|--------|-------|
| Test files created | 13 |
| Total tests | 1,203 |
| Test failures | 0 |
| Commits | 30 |
| Fix rounds across all tasks | 8 |
| Quarantined example scripts | 5 groups |
| Release binaries | 8 (all pass) |
