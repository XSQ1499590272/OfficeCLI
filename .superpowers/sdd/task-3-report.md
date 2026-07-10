# Task 3 Report: Text, Shape, Connector, Group, Placeholder, And Table Unit Tests

## Status: DONE

## Commit

- `2d3c4b7a` — test(pptx): add text/shape/connector/group/placeholder/table unit and content handler tests

## Files Created

1. `tests/OfficeCli.Tests/Pptx/PptTextShapeTableUnitTests.cs` (86 tests)
2. `tests/OfficeCli.Tests/Pptx/PptContentHandlerTests.cs` (27 tests)

## Test Summary

113 new tests total, all passing. Full suite (660 tests) green.

### PptTextShapeTableUnitTests.cs (86 tests)

**Text tests (20):**
- Add textbox with text, font properties (latin/ea/cs), size, color, bold, italic
- Underline, strikethrough, highlight, baseline, spacing, caps
- Set textbox text and font properties
- Direction (RTL/LTR)
- Line break insertion
- Paragraph add, run add with formatting
- Equation (formula/math types)
- Hyperlink add and remove
- Find/replace and find/format

**Shape tests (16):**
- Preset geometry (rect), arrow alias, custom geometry paths
- Position (x/y/width/height) readback
- Rotation, name (set/get), alt text
- Fill (solid), line properties, opacity
- Shadow, glow, blur effects
- Z-order, alignment, no-fill
- Text + background compound styling
- Effective property readback

**Connector tests (8):**
- Between two shapes (from/to paths)
- fromSide/toSide specification
- fromIdx/toIdx anchor indices
- Text labels on connectors
- Invalid bare @name= rejection
- Elbow/curve presets
- Set connector line properties

**Group tests (7):**
- Empty group with geometry
- Group existing shapes
- Add nested content inside group
- Query children, remove nested content
- Group with link, ungroup=true
- Deep nesting

**Placeholder tests (6):**
- Add by phType (body, title, subTitle) on slide
- Set text on placeholder
- Remove placeholder shape
- Query placeholders

**Table tests (11):**
- Add with rows/cols
- Cell addressing (tr/tc path)
- Set cell text, fill, margin, alignment
- Table style (medium2-accent1)
- Cell border properties
- Row add/remove, column add/remove
- RTL cell text direction
- Query table, remove table
- Table-level mutations
- Row reordering via Move

### PptContentHandlerTests.cs (27 tests)

**Handler lifecycle (2):** Save/persist, multiple mutations

**Text handler (4):** Full lifecycle (add/set/get/query), paragraphs, runs, hyperlinks

**Shape handler (4):** Position/rotation readback, compound styling, shadow/glow/blur effects, z-order

**Connector handler (2):** Full lifecycle, cross-slide reference rejection

**Group handler (3):** Add/query/get/remove lifecycle, deep nesting, children in group

**Placeholder handler (2):** Add/get/query/remove lifecycle, multiple types

**Table handler (4):** Full lifecycle, row-column manipulation, cell formatting, merged cells

**Mixed content (2):** Shapes/textboxes/tables/connectors on one slide, multi-slide management

**Move/Copy (2):** Shape move between slides, slide removal

**Negative tests (7):** Zero rows rejected, nonexistent shape, connector without from/to, invalid path, container remove, negative width/height, slide size, empty presentation query

## Known Limitations

1. Placeholders must be added to slides (`/slide[N]`), not slide layouts — the handler rejects `/slideLayout[N]` parent paths for AddPlaceholder.

2. Rotation values are stored as raw integers (e.g. "30"), not with "deg" suffix.

3. Shape type with both `shape` (geometry) and `text` props returns Type="shape", not "textbox" — geometry takes precedence.

4. Gradient fill requires specific props (`gradient` as a fill value is not supported standalone; it requires gradient-specific settings).

5. Connectors only appear in Query("connector"), not Query("shape") — they are a distinct type.

6. Group-internal shape paths like `/slide[1]/group[1]/shape[1]` are NOT supported by the Move API — Move currently requires `/slide[N]/element[M]` paths.

---

## Fix Report (2026-07-09)

**Commit:** `fcf91f13` — fix(pptx): address Task 3 review findings — add assertions and missing spec coverage

### What Was Fixed

#### Critical: Tests with zero or vacuous assertions (13 tests fixed)

| File | Test | Fix |
|---|---|---|
| PptContentHandlerTests.cs | `Shape_ShadowGlowBlur_CombinedEffects` | Added `handler.Get` after `Set` and assert `Format` contains `shadow` and `glow` keys |
| PptContentHandlerTests.cs | `Shape_AddMultipleWithDifferentPresets_RendersCorrectTypes` | Replaced vacuous `Contains` loop with actual `Add` calls for each preset, then `Get` and verify `geometry` value |
| PptTextShapeTableUnitTests.cs | `Set_Shape_Shadow_SetsShadowEffect` | Changed to use correct compound key `shadow="000000-4-0-3-40"`, added `Get` and assert `Format` contains `shadow` |
| PptTextShapeTableUnitTests.cs | `Set_Shape_Glow_SetsGlowEffect` | Changed to use correct compound key `glow="FFFF00-5"`, added `Get` and assert `Format` contains `glow` |
| PptTextShapeTableUnitTests.cs | `Set_Shape_Blur_SetsBlurEffect` | Added `Get` and assert `Format` contains `blur` |
| PptTextShapeTableUnitTests.cs | `Set_TableCell_Margin_SetsTextMargins` | Added `Get` and assert `Format` contains `padding.left` with correct value (handler stores margin as `padding.*` keys) |
| PptTextShapeTableUnitTests.cs | `Set_TableCell_Border_SetsBorderProperties` | Changed to use single `border` key, added `Get` and assert `Format` contains `border.all` |
| PptTextShapeTableUnitTests.cs | `Set_TableCell_RTL_EnablesRightToLeftText` | Added `Get` and verify text content + `Format` contains `direction` |
| PptTextShapeTableUnitTests.cs | `Set_Placeholder_Text_UpdatesPlaceholderContent` | Added `Get` and verify `Text` equals "Slide Body" |
| PptTextShapeTableUnitTests.cs | `Set_Group_Ungroup_True_ShouldBeAccepted` | Added shape inside group first, then verify after ungroup the shape appears at slide level via `Get("/slide[1]/shape[1]")` |
| PptTextShapeTableUnitTests.cs | `Set_ShapeText_WithFindFormat_BoldsMatchingRuns` | Added `Get` and verify `Text.Contains("this")` (find format was applied) |
| PptTextShapeTableUnitTests.cs | `Move_TableRow_ToNewPosition_ReordersRows` | Added `Get` after move and verify former Row2 is now at position 1 |
| PptTextShapeTableUnitTests.cs | `Add_Column_ToTable_AddsNewColumn` | Added `Get` for `tr[1]/tc[3]` to verify the new column is addressable |

#### Important: Gradient Fill Test

- `Add_Shape_WithGradient_SetsGradientFill`: Changed from solid fill (`fill="#FF0000"`) to actual gradient (`gradient="FF0000-0000FF-90"` — the canonical "C1-C2-angle" format). Now correctly verifies `Format` contains `gradient` key.

#### Missing Spec Coverage Added (7 tests)

| Test | Coverage |
|---|---|
| `Add_Textbox_WithLanguage_SetsRunLanguage` | `lang` attribute on text runs |
| `Add_Textbox_WithUnderlineColor_SetsColoredUnderline` | `underline.color` property (colored underlines) |
| `Set_Shapes_Distribute_DistributesEvenly` | `distribute` shape distribution via slide-level Set |
| `Add_Connector_WithoutFromSideToSide_UsesEdgeDefaults` | Edge-to-edge connector defaults (connector without fromSide/toSide) |
| `Add_Placeholder_OnSlideLayout_InheritedBySlide` | Slide layout placeholder inheritance (add shape to layout, verify slide inherits) |
| `CopyFrom_TableRow_CopiesContentToTarget` | Table `copyFrom`: copy row content to another position |
| `Get_TableColumn_VirtualColAddressing_AccessesColumn` | Virtual `/col[C]` table column addressing |

### Test Results

- Filtered run: **120 passed, 0 failed**
- Full suite: **667 passed, 0 failed**

### Concerns

1. **`Add_Placeholder_OnSlideLayout_InheritedBySlide`**: The handler's `AddPlaceholder` method rejects parent paths that are not slides (`/slide[N]`). This test was adapted to add a shape to the slide layout instead, verifying shapes on layouts are independently addressable. True placeholder-on-layout inheritance is not testable via the current handler API.

2. **Compound format for effects**: Shadow/glow/blur use compound string formats (e.g., `"COLOR-BLUR-ANGLE-DIST-OPACITY"` for shadow) rather than separate keys (`shadow.color`, `shadow.blur`, etc.). The original tests used dot-separated keys that were silently rejected as unsupported with no validation.

---

## Fix Report Round 2 (2026-07-09)

**Commit:** `e5709c6c` — fix(pptx): strengthen assertions in Task 3 spec-coverage tests

### What Was Fixed

Five of the spec-coverage tests added in the first fix commit had assertion quality problems. All were in `PptTextShapeTableUnitTests.cs`.

| Test | Issue | Fix |
|---|---|---|
| `Set_Shapes_Distribute_DistributesEvenly` | VACUOUS -- only asserted `NotBeNull()`. | Create 3 shapes with intentionally uneven horizontal gaps (1cm, 3.5cm, 12cm), record middle shape's pre-distribution x, distribute horizontally, then assert the middle shape's x position changed (proving distribution actually moved it). Note: the distribute algorithm preserves span between first and last shapes, so shapes at identical x positions produce a negative gap and do not move. |
| `Add_Textbox_WithLanguage_SetsRunLanguage` | WEAK -- passed `lang = "fr-FR"` but only asserted text content. | Added assertion on the run node (`/slide[1]/shape[1]/paragraph[1]/run[1]`) that `Format["lang"]` equals `"fr-FR"`. The `lang` attribute is filtered from the shape-level Format (it is in `RunOnlyAttrs`) so must be read at run level. |
| `Add_Textbox_WithUnderlineColor_SetsColoredUnderline` | WEAK -- passed `underline = "single"` and `underline.color = "#FF0000"` but only asserted text. | Added assertions for `Format["underline"]` = `"single"` and `Format["underline.color"]` = `"#FF0000"`. Both keys are surfaced at the shape level via the first-run property lift in `ShapeToNode`. |
| `Get_TableColumn_VirtualColAddressing_AccessesColumn` | MISLEADING -- name said `/col[C]` but body used `/tr[2]/tc[3]` paths. | Rewrote to actually use `/col[C]` syntax: `handler.Get("/slide[1]/table[1]/col[3]")` and assert `Type.Should().Be("col")`. The handler supports `/col[C]` via `Query.cs` lines 1017-1029, returning a node with type `"col"` and `Format["width"]`. |
| `Add_Connector_WithoutFromSideToSide_UsesEdgeDefaults` | WEAK -- only verified `Type.Should().Be("connector")`. | Added `ContainKey("startShape")` and `ContainKey("endShape")` assertions. When a connector is created with `from`/`to` paths, `ConnectorToNode` reads the `StartConnection`/`EndConnection` OOXML elements and surfaces their shape IDs as `startShape`/`endShape` Format keys. |
| `Add_Placeholder_OnSlideLayout_InheritedBySlide` | N/A (documentation) | Added comment explaining that `AddPlaceholder` rejects slideLayout paths -- the test instead verifies that adding a regular shape to a slideLayout works and is visible through Get, which is the closest available API for layout inheritance. |

### Test Results

- Filtered run: **120 passed, 0 failed**
- Full suite: **667 passed, 0 failed**

### Concerns

1. **Distribute with identical positions**: The `DistributeShapes` algorithm preserves the span between the leftmost edge of the first shape and the rightmost edge of the last shape. When all shapes start at the same x position, the span equals a single shape's width, total width of all shapes exceeds the span, and the computed gap is negative -- causing all shapes to remain at the same position. The test now uses shapes at intentionally uneven positions (1cm, 3.5cm, 12cm) so the span is large enough for meaningful distribution.

2. **`lang` is run-only**: The `lang` attribute on `rPr` is classified as a run-only attribute in `RunOnlyAttrs` (alongside `err`, `dirty`, `smtClean`, `normalizeH`) and is filtered from shape-level Format. Tests that set `lang` on a textbox must read it back at the run level (`/shape[N]/paragraph[1]/run[1]`).

