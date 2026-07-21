# Pitch Deck Agent Skill

Author: xiesq, 2026-07-15

## ⚠️ Help 优先规则

**This skill teaches what a fundraising deck requires, not every command flag.** When a prop name, enum value, or preset is uncertain, consult help BEFORE guessing.

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "pptx"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "pptx",
      "chart"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "pptx",
      "connector"
    ]
  }
}
```

Help reflects the installed version. When this skill and help disagree, **help wins.**

Load the base skill first:

```json
{
  "tool": "{{OFFICE_LOAD_SKILL_TOOL}}",
  "arguments": {
    "command_arguments": [
      "pptx"
    ]
  }
}
```

`{{OFFICE_LOAD_SKILL_TOOL}}` only reads guidance; it does not install or modify local skills.

## Mental Model & Inheritance

**This skill is a scene layer on top of `pptx`.** Inherit visual floor, 12-column grid (33.87×19.05cm), palettes, chart-choice, connector canon (`tailEnd=triangle`), Delivery Gate 1–5a. This file adds fundraising deltas: stage diagnosis, 5 赛道 arcs, key-slide recipes, numbers convention, VC ship-check, Gate 6. Base rules → `see pptx §X`. Load `pptx` first.

### Data truth rule

Organize user-provided metrics into the narrative. **Do not invent** traction, TAM/SAM/SOM, revenue, customer counts, team credentials, NCT, ORR, or benchmarks. Unknown numbers = explicit gap / ask the user. Example JSON numbers are **structural skeletons only** — replace with user-provided or sourceable data before delivery.

## 宿主工具与执行规范

Agent 通过宿主 Tool JSON 调用 Office 工具，不使用 shell / bash / heredoc。

### Run vs Batch

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- Batch 后必须独立用 Run 做 readback / `view` / `validate`。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。
- props 仅允许 string / number / bool。

已通过 inspect 确认前三张 slide 存在，且业务接受部分成功或正在可丢弃副本上工作时，可将三个互不依赖的背景 mutation 合并为 Batch：

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/pitch-deck.pptx",
    "operations": [
      {"command":"set","path":"/slide[1]","props":{"background":"0B1020"}},
      {"command":"set","path":"/slide[2]","props":{"background":"111827"}},
      {"command":"set","path":"/slide[3]","props":{"background":"0F172A"}}
    ],
    "stop_on_error": true
  }
}
```

逐项检查结果后，再用独立 Run 完成 readback、visual view 与 validate。

### 文本语义

Path 含 `[N]` 原样传入 Tool 参数。货币金额中的 `$`（如 `$35M`、`$1.2B TAM`）直接写在 JSON 字符串内（无 shell 展开）。写入后 Run `view text` 核对 `$` 是否保留——这是 pitch deck 头号文本失败模式（cover / ask / financials / milestones）。

## What "pitch deck" means here (identity)

A pitch deck is a pptx with a **fundraising layer**: VC narrative arc, verifiable metrics, stage-appropriate density, founder credibility. ~3 seconds per slide; **every slide carries one investable proposition**. Six deltas:

1. **Stage determines everything.** Series A / B / C each dictates slide count, narrative weight, which metrics are must-haves, and tolerance for unit-econ sophistication.
2. **Narrative arc beats feature dump.** 10 essential slides in a fixed order: cover → problem → solution → market → product → model → traction → team → financials → ask.
3. **Numbers are a contract.** TAM/SAM/SOM must be clean three-layer; CAC/LTV must have a payback line; ARR ≠ revenue; Use-of-Funds must be a four-bucket pie.
4. **Team slide carries prior companies.** Avatar grid alone reads as a student project.
5. **Traction chart y-axis starts at 0.** A "hockey stick" starting at `y_min = 80% of current` is a visual lie.
6. **The ask is a slide, not a footnote.** `$XX M` hero + four-bucket Use-of-Funds + runway length.

### Reverse handoff — when to go BACK to pptx base

Stay in **pptx base** for board reviews, all-hands, sales, launches, training. Use **this skill** when: (a) specific round or VC meeting, AND (b) ≥4 of {problem, traction, team credentials, Use-of-Funds, unit econ, financials}. Corporate BU quarterly ask → pptx; bridge round framed as "board review" → here.

## Series A / B / C stage diagnosis (decision tool)

**Read this before writing a single mutation.** Pick the row that matches the user's description — everything downstream derives from this one call.

| Stage | Revenue band | Team | Slide count | Dominant narrative (weight) | Must-have data | Common red flag |
|---|---|---|---|---|---|---|
| **Seed** | $0 – $1M ARR (often pre-rev) | 2 – 8 FTE | 10 – 12 | Problem (30%) + Solution (25%) + Team (15%) + Market (15%) + Traction (15%) | Founder-market fit story; 1 – 2 design-partner / pilot logos; top-down TAM ok | Over-claiming traction (10 customers = "market proven") |
| **Series A** | $1 – $5M ARR | 10 – 25 FTE | 12 – 16 | Problem (20%) + Solution (20%) + **Market "why now"** (15%) + Product (15%) + Traction (20%) + Team (10%) | PMF proof (NRR > 110%, low churn), bottom-up TAM/SAM, pipeline / pilots converted | Bottom-up TAM feels fabricated; CAC not yet meaningful but shown anyway |
| **Series B** | $5 – $30M ARR | 30 – 100 FTE | 18 – 22 | **Traction + Unit econ (30%)** + Market + Product + Team + Financials (ask) | ARR curve starting at 0; NRR, CAC, LTV, payback (< 18 mo ideal); cohort retention; logo wall | No unit-econ slide; CAC payback > 24mo without explanation; Use-of-Funds missing % |
| **Series C** | $30M+ ARR | 100+ FTE | 20 – 24 | **Financials + Scale + Moat (40%)** + Market expansion + Team depth | Multi-year GAAP, rule-of-40, GM trajectory, international expansion plan, defensibility | No moat slide; revenue growth without margin story; team slide has no prior CEO / CFO |
| **Bridge / SAFE** | any | any | 8 – 10 | **Specific bridge reason** + runway math + commitments | Prior round context; specific milestone the bridge funds; committed investor amount | Treating a bridge like a Series A — too many slides dilutes the ask |

**Decision procedure.** From one or two user sentences ("Series B, $18M ARR, 120 customers, $35M raise"), pick exactly one stage row.

**Corner cases.** Bridge rounds & convertibles between A → B are closer to A or B depending on whether the bridge milestone is "finish PMF" (A shape) or "hit unit-econ target" (B shape). "Extension" rounds at the same stage reuse the earlier stage's skeleton and add a one-slide "progress since last round" update.

**Non-SaaS stage overrides** (substitute Series B unit-econ + Gate 6.3):

| Vertical | Revenue "band" at Series B | "Unit econ" equivalent | Gate 6.3 substitute |
|---|---|---|---|
| **Bio / Clinical-stage** | pre-rev, 20–60 FTE | burn rate + runway to next milestone (IND / Ph1 readout / BLA) | `shape:contains("ORR")` OR `contains("Pipeline")` OR `contains("BLA")` OR `contains("runway")` ≥ 1 |
| **Deep Tech / Frontier** | pre-rev or early pilot rev | technical milestones + TRL level + benchmark vs SoTA | `shape:contains("TRL")` OR `contains("benchmark")` ≥ 1 |
| **Marketplace / Network** | GMV $10–100M | take rate + cohort retention + liquidity | `shape:contains("GMV")` + `contains("take rate")` ≥ 1 |
| **Consumer hardware** | $2–15M revenue (shipped units) | contribution margin + repeat rate + blended CAC | `shape:contains("repeat")` OR `contains("contribution")` ≥ 1 |

Bio Series B: burn + runway-to-milestone IS the unit-econ story.

## 赛道 arc templates (5 families)

5 mainstream verticals. Pick the vertical row; the slide skeleton is a starting point. Slide counts assume the matching stage row above.

### (1) B2B SaaS / Enterprise software

Canonical arc — Series B example (20 slides): cover · TL;DR · problem · problem evidence · solution · product loop · market TAM/SAM/SOM · **unit economics (CAC / LTV / payback / GM)** · ARR trajectory · retention cohort · logo wall · team · competitors · financials 4-year · ask. Must-have: unit-econ slide from Series A onward; logo wall from Series B onward.

### (2) Consumer (B2C app / consumer hardware / D2C)

Narrative-driven. Series A example (14 slides): cover · hook · problem (lived experience) · solution (product shots) · product-experience flow · "why now" market window · pre-order / crowdfunding / early-sales evidence · retention / engagement (DAU, D30) · market · competitive positioning · founder story + team · press / endorsements · financials · ask. Must-have: product visuals on ≥ 3 slides; "why now" slide; engagement metric not just revenue.

### (3) Deep Tech / Frontier tech

Technology credibility is the sell. Series B example (22 slides): cover · thesis · problem (current state of art) · solution (technical approach) · **technology architecture** · benchmarks vs SoTA · pipeline / TRL levels · market · business model · early commercial traction · IP / patents · team · partners · financials · ask. Must-have: benchmark slide; IP slide; team slide dense with PhDs / prior-lab names.

### (4) Marketplace / Network business

Liquidity is the metric. Series A example (15 slides): cover · problem · solution · product demo (both sides) · network effects diagram · early liquidity · cohort retention · geographic / category expansion · competitive positioning · take-rate model · team · financials · ask. Must-have: liquidity metric slide; cohort retention chart; network-effect diagram.

### (5) Bio / Life sciences / Healthtech

Regulatory pipeline IS the business. Series B example (22 slides): cover · unmet medical need · scientific rationale · preclinical / clinical data · **pipeline chart** · differentiation vs standard of care · IP / exclusivity · regulatory strategy · market · commercial strategy · partnerships · team (CSO / CMO with prior FDA wins) · financials (burn to next milestone) · ask. Must-have: pipeline chart; clinical data slide; team slide with prior regulatory wins.

**Cross-vertical rule.** You can mix elements across templates, but never drop a must-have from your primary vertical.

## Slide Patterns (layout canon)

Patterns are **layout geometry**; recipes below are **narrative intent**. Pick the pattern first, then fill it with recipe content.

**Speaker notes rule.** Every content slide (non-cover, non-closing) MUST carry speaker notes via `add` `type=notes`. Missing notes = not shippable — inherits pptx H7. Confirm prop names with Help pptx notes.

**Pattern reuse discipline.** Never run the same pattern on two consecutive slides. Alternate C.2 with C.4 or C.5b to break rhythm.

**Vertical centering.** When a slide carries fewer elements than the pattern's maximum, nudge y-positions down 2–3cm to center the visual weight. Tables below assume full content.

### C.1 Title / Cover (dark gradient)

3–4 text shapes on a gradient fill. Slide 1 in every deck.

| Element | X | Y | Width | Height | Font / size |
|---|---|---|---|---|---|
| Title | 2cm | 5cm | 29.87cm | 4cm | serif bold, ≥ 36pt (44 typical) |
| Tagline | 2cm | 10cm | 29.87cm | 2cm | sans 18–22pt |
| Meta (round · $ · date) | 2cm | 13cm | 29.87cm | 1.5cm | sans 12–16pt |

**Use:** Cover recipe 1. Dark gradient bg (e.g. `1E2761 → 0D1F35`). Title wraps → add height, never drop below 36pt. Transition: fade.

### C.2 3-Stat callout row

Title + 3 big-number / label pairs across. Default for Problem / Why-Now / Traction-callout.

| Element | X | Y | Width | Height | Font / size |
|---|---|---|---|---|---|
| Title | 1.5cm | 1cm | 30.87cm | 3cm | serif bold ≥ 36pt |
| Stat 1 number | 2cm | 5cm | 9cm | 4cm | serif bold 60–64pt |
| Stat 1 label | 2cm | 9.5cm | 9cm | 2cm | sans ≥ 16pt (H4 floor) |
| Stat 2 number / label | 12.5cm | (same) | 9cm | (same) | (same) |
| Stat 3 number / label | 23cm | (same) | 9cm | (same) | (same) |

### C.3 4-Stat callout row

Same geometry as C.2 but 4 columns. Numbers 60pt, width 7cm each. X positions: `1.5 / 9.5 / 17.5 / 25.5cm`. Prefer C.2 if in doubt.

> **Wrap warning.** At 60pt in 7cm width, dollar patterns with both `$` and `.` fail: `$9.4M` wraps. Safe: `$9M`, `$96B`, `$4K`. Values ≥ 6 chars will wrap — drop font to 44–48pt, abbreviate, or shift to C.2.

### C.4 Chart + Context (chart left, stats right)

Chart left 55%, 2–3 stacked callouts on the right. Default for Traction / Financials / Market-sizing-with-context.

| Element | X | Y | Width | Height |
|---|---|---|---|---|
| Title | 2cm | 1cm | 29.87cm | 3cm |
| Chart | 2cm | 4cm | 17cm | 13cm |
| Stats column | 21cm | 4cm+ | 11cm | ~3.7cm per pair |

Post-batch for column/bar charts: Run `set` on the chart with `gap=80` to tighten bar spacing.

### C.5 Icon-in-circle grid (3-row vertical)

Icon circles at y `4.5 / 8.5 / 12.5cm` (2.5cm ellipse); labels at x=5.5cm. Choose C.5b when exactly 4 parallel items.

### C.5b 2×2 Feature grid (4 parallel items)

| Element | X | Y | Width | Height |
|---|---|---|---|---|
| Slide title | 2cm | 1cm | 29.87cm | 2.5cm |
| Card 1 bg | 1.5cm | 4cm | 14.5cm | 7cm |
| Card 2 bg | 17.5cm | 4cm | 14.5cm | 7cm |
| Card 3 bg | 1.5cm | 12cm | 14.5cm | 7cm |
| Card 4 bg | 17.5cm | 12cm | 14.5cm | 7cm |

**Z-order canon.** Each card's `roundRect` background must be added immediately before that card's icon / title / body — insertion order paints. Per-card sequence `bg → ellipse → title → body`. Dark-background variant: card fill `1A2540`, body text `FFFFFF` / `E8E8E8`. Write literal hex in JSON.

## Key-slide recipes (10 essentials)

The 10 slides every pitch deck carries. Each recipe: **visual outcome** + **one Tool JSON conversion of the CLI example** + **QA one-liner**. Recipes reference Slide Patterns above. Path examples use `/workspace/pitch.pptx`. Example numbers are skeletons — **Do not invent** real metrics.

**Long-title wrap rule.** A 36pt+ title that wraps to 2 lines: add `height` — never drop the font below 36pt.

> **Chart `series1.color=` on `add` works.** Verify with a readback `get` on `/slide[N]/chart[1]/series[1]` if needed.

Lifecycle:

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": [
      "/workspace/pitch.pptx"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "open",
    "command_arguments": [
      "/workspace/pitch.pptx"
    ]
  }
}
```

### (1) Cover slide — company · tagline · round · date

**Visual outcome.** Dark navy fill, centered 44pt company name, 20pt tagline, 16pt meta (round + amount + date), thin brand band at bottom (0.5cm, accent).

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=1E2761"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[1]",
      "--type",
      "shape",
      "--prop",
      "name=BrandBand",
      "--prop",
      "geometry=rect",
      "--prop",
      "fill=CADCFC",
      "--prop",
      "x=0cm",
      "--prop",
      "y=18.5cm",
      "--prop",
      "width=33.87cm",
      "--prop",
      "height=0.55cm"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[1]",
      "--type",
      "shape",
      "--prop",
      "name=CoverTitle",
      "--prop",
      "text=Acme DevOps",
      "--prop",
      "x=2cm",
      "--prop",
      "y=7cm",
      "--prop",
      "width=29.87cm",
      "--prop",
      "height=3cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=44",
      "--prop",
      "bold=true",
      "--prop",
      "color=FFFFFF",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[1]",
      "--type",
      "shape",
      "--prop",
      "name=Tagline",
      "--prop",
      "text=Kubernetes observability, built for production at scale",
      "--prop",
      "x=2cm",
      "--prop",
      "y=10.5cm",
      "--prop",
      "width=29.87cm",
      "--prop",
      "height=1.5cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=20",
      "--prop",
      "color=CADCFC",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[1]",
      "--type",
      "shape",
      "--prop",
      "name=CoverMeta",
      "--prop",
      "text=Series B · $35M · April 2026",
      "--prop",
      "x=2cm",
      "--prop",
      "y=15cm",
      "--prop",
      "width=29.87cm",
      "--prop",
      "height=1.2cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=16",
      "--prop",
      "color=FFFFFF",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

**QA.** Cover has 4 discrete elements (brand band + title + tagline + meta). 80%-whitespace covers fail the pptx "cover ≥ 60% filled" floor.

**Consumer variant.** Consumer decks may add a single dominant motif — hero product shot, oversized company name (60–96pt), or symbolic mark. Keep tagline + round + date identical. SaaS / B2B may skip.

### (2) Problem slide — industry pain in 1 sentence + 3 data cards

**Visual outcome.** 36pt title stating the pain (not "The Problem"). Below, three equal-width data cards: giant number (60pt) + qualifier (18pt) + source footnote (12pt gray).

Grid math for 3 cards, 1.5cm margins, 0.76cm gap: `usable = 33.87 − 3 − 2·0.76 = 29.35`, `col_width = 29.35 / 3 = 9.78cm`. x-positions: `1.5 / 12.04 / 22.58`.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[2]",
      "--type",
      "shape",
      "--prop",
      "text=Kubernetes debugging burns 12 engineering hours / incident",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2.5cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[2]",
      "--type",
      "shape",
      "--prop",
      "name=PC1",
      "--prop",
      "geometry=roundRect",
      "--prop",
      "fill=F5F7FA",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=5cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=10cm"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[2]",
      "--type",
      "shape",
      "--prop",
      "text=73%",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=6cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=3cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=60",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[2]",
      "--type",
      "shape",
      "--prop",
      "text=of incidents take > 1 hour to diagnose",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=9.5cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=3cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=18",
      "--prop",
      "color=333333",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[2]",
      "--type",
      "shape",
      "--prop",
      "text=Source: 2025 DORA Report",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=13cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=1cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=12",
      "--prop",
      "italic=true",
      "--prop",
      "color=666666",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

Repeat the 4-block card pattern at x=`12.04cm` and x=`22.58cm` for cards 2 and 3.

**QA.** Run `query` for `shape:contains("Source")` — expect ≥ 3. If zero sources, VCs will not trust a single number.

### (2b) Why Now slide — Consumer / Seed / early A must-have

**Visual outcome.** 3 cards across: each = **trigger headline** (24pt bold) + **data point** (60pt) + **one-line implication** (16pt) + **source footnote** (12pt gray). Reuse Problem grid (`col=9.78cm`, x = `1.5 / 12.04 / 22.58`).

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[3]",
      "--type",
      "shape",
      "--prop",
      "text=Why now: three converging triggers",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2.5cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[3]",
      "--type",
      "shape",
      "--prop",
      "geometry=roundRect",
      "--prop",
      "fill=F5F7FA",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=5cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=10cm"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[3]",
      "--type",
      "shape",
      "--prop",
      "text=BOM cost",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=5.5cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=1.2cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=24",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[3]",
      "--type",
      "shape",
      "--prop",
      "text=−90%",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=7cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=3cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=60",
      "--prop",
      "bold=true",
      "--prop",
      "color=B85042",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[3]",
      "--type",
      "shape",
      "--prop",
      "text=Wearable BOM fell 90% since 2021; sub-$40 retail now viable",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=11cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=16",
      "--prop",
      "color=333333",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[3]",
      "--type",
      "shape",
      "--prop",
      "text=Source: IDC Wearables Teardown 2025",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=13.5cm",
      "--prop",
      "width=9.78cm",
      "--prop",
      "height=1cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=12",
      "--prop",
      "italic=true",
      "--prop",
      "color=666666",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

Repeat the card pattern at x=`12.04cm` and x=`22.58cm`. Card 2 pattern (prose): Oura IPO 2024 / +$2.4B valuation / category proven. Card 3: On-device LLM (Llama 3.2) / Q4-24 / privacy moat viable.

**QA.** 3 cards, each with a date/year citation, each card ≤ 30 words. Query `shape:contains("2024")` + `shape:contains("2025")` ≥ 2 combined.

### (3) Solution slide — product in one sentence + 3-step "how it works"

**Visual outcome.** 36pt title naming the product pattern (not "Our Solution"). Below: 3 or 4 rounded boxes horizontally at y=7cm with elbow connectors + triangle arrowheads. Reuse pptx Recipe (c) flowchart — orchestration, not a new primitive.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[4]",
      "--type",
      "shape",
      "--prop",
      "name=SolTitle",
      "--prop",
      "text=Correlate K8s events across 3 data planes in 90 seconds",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2.2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=32",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[4]",
      "--type",
      "notes",
      "--prop",
      "text=Walk the three-step mechanism; do not pitch brand slogans."
    ]
  }
}
```

3 boxes across: gap = `(33.87 − 3 − 3·7) / 2 = 4.93cm`; x = `1.5, 13.43, 25.36`. Connectors + arrowheads: `tailEnd=triangle` ALWAYS. Full flowchart batch → see pptx §Creating and Editing (c) 4-step flowchart; swap N from 4 boxes to 3.

**Product-pattern title rule.** Verb + differentiated mechanism + metric. "Correlate K8s events across 3 data planes in 90 seconds" is specific; "The future of observability" is not.

**QA.** Count connectors ≥ (step_count − 1). Every connector must have `tailEnd=triangle`. Title ≤ 12 words.

### (4) Market slide — TAM / SAM / SOM nested columns

**Visual outcome.** 36pt title "Market: $X.YB growing Z% CAGR". Below: bar/column chart for TAM/SAM/SOM. Bottom footnote cites **top-down vs bottom-up source** — pick one methodology per deck.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[5]",
      "--type",
      "shape",
      "--prop",
      "text=$42B observability market, 18% CAGR",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[5]",
      "--type",
      "chart",
      "--prop",
      "chartType=bar",
      "--prop",
      "series1.name=USD (billions)",
      "--prop",
      "series1.values=42,8.4,0.62",
      "--prop",
      "series1.color=1E2761",
      "--prop",
      "categories=TAM,SAM,SOM (5-yr)",
      "--prop",
      "x=2cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=22cm",
      "--prop",
      "height=12cm",
      "--prop",
      "title=Market sizing — bottom-up by enterprise count × ACV"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[5]",
      "--type",
      "shape",
      "--prop",
      "text=Source: Gartner 2025 APM Magic Quadrant; SAM = 20% of TAM (K8s-first shops); SOM = 7.4% of SAM over 5 years at 18-24% share.",
      "--prop",
      "x=2cm",
      "--prop",
      "y=16.5cm",
      "--prop",
      "width=29.87cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=12",
      "--prop",
      "italic=true",
      "--prop",
      "color=666666",
      "--prop",
      "fill=none"
    ]
  }
}
```

**QA.** Top-down vs bottom-up MUST be declared in the source footnote.

### (5) Product slide — screenshot + 3 bullets OR 3-card feature grid

**Visual outcome.** (a) hero screenshot left 60% + 3 one-line feature bullets right (≥ 18pt). (b) 3 feature cards. Pick (a) for consumer / app, (b) for B2B / infrastructure.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "picture",
      "--prop",
      "src=/workspace/assets/product_hero.png",
      "--prop",
      "x=1cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=18cm",
      "--prop",
      "height=13cm"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]/picture[1]",
      "--prop",
      "alt=Product UI: dashboard with 12 K8s clusters, live correlation graph"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "shape",
      "--prop",
      "text=Auto-correlate across 3 data planes",
      "--prop",
      "x=20cm",
      "--prop",
      "y=5cm",
      "--prop",
      "width=12cm",
      "--prop",
      "height=1.5cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=20",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

Repeat for bullets 2 and 3 at y=`7.5cm` / y=`10cm`.

**QA.** Picture alt present (`query 'picture:no-alt'` empty). Bullets each ≥ 18pt. No "Lorem" / "product name here" / `{{...}}` tokens.

### (6) Business model slide — unit econ or revenue model

**Visual outcome.** Decision tree by vertical:
- **SaaS / Enterprise (Series A+)** — 4 KPI callouts: CAC / LTV / Payback / GM (reuse pptx Recipe (e)).
- **Consumer / D2C** — AOV · repeat-purchase rate · contribution margin · blended CAC.
- **Marketplace** — GMV / take-rate / liquidity / cohort retention.
- **Bio / Deep tech** — revenue model (license / milestone / royalty) with assumed ranges.

Title names the dominant metric (e.g. "LTV:CAC 4.7x · 14-month payback · 78% gross margin"), not "Business Model".

SaaS card values (prose skeleton — full 4-card batch → pptx §(e); adapt card count 3→4 and width 9.78cm→7.15cm):
- Card 1 (LTV): `$420K` / "Lifetime value" / "floor: ARPU × GM / churn"
- Card 2 (CAC): `$90K` / "Acquisition cost" / "fully-loaded S&M spend"
- Card 3 (Payback): `14 mo` / "CAC payback" / "VC floor: < 18 mo"
- Card 4 (GM): `78%` / "Gross margin" / "SaaS floor: 70%+"
- Grid for 4 cards: `usable = 33.87 − 3 − 3·0.76 = 28.59`, `col = 7.15cm`

**QA.** For Series B+, all four of {CAC, LTV, payback, GM} present via `query` `shape:contains(...)`.

### (7) Traction slide — ARR curve that starts at 0

**Visual outcome.** Line chart ~60% width; ARR y-axis **starting at 0**. Right-side callout: current ARR + growth + NRR. Series B+: second row cohort / logo wall.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[7]",
      "--type",
      "shape",
      "--prop",
      "text=ARR: $0 → $18M in 24 months",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[7]",
      "--type",
      "chart",
      "--prop",
      "chartType=line",
      "--prop",
      "series1.name=ARR",
      "--prop",
      "series1.values=0.2,0.6,1.4,3.2,6.1,11.3,15.8,18.0",
      "--prop",
      "series1.color=1E2761",
      "--prop",
      "categories=Q1-24,Q2-24,Q3-24,Q4-24,Q1-25,Q2-25,Q3-25,Q4-25",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=21cm",
      "--prop",
      "height=13cm",
      "--prop",
      "title=Quarterly ARR ($M) — y-axis anchored at 0",
      "--prop",
      "axismin=0"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[7]",
      "--type",
      "shape",
      "--prop",
      "geometry=roundRect",
      "--prop",
      "fill=1E2761",
      "--prop",
      "line=none",
      "--prop",
      "x=23.5cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=8.8cm",
      "--prop",
      "height=13cm"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[7]",
      "--type",
      "shape",
      "--prop",
      "text=$18M",
      "--prop",
      "x=23.5cm",
      "--prop",
      "y=5cm",
      "--prop",
      "width=8.8cm",
      "--prop",
      "height=3cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=64",
      "--prop",
      "bold=true",
      "--prop",
      "color=FFFFFF",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[7]",
      "--type",
      "shape",
      "--prop",
      "text=ARR · +312% YoY · NRR 128%",
      "--prop",
      "x=23.5cm",
      "--prop",
      "y=9cm",
      "--prop",
      "width=8.8cm",
      "--prop",
      "height=3cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=18",
      "--prop",
      "color=CADCFC",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

**`axismin=0` is load-bearing** — without it, auto-scale creates the hockey-stick lie. Gate 6 greps this.

**QA.** Run `get` on the chart; readback `axisMin` / `axismin` returns `0`.

### (8) Team slide — avatars + names + prior companies

**Visual outcome.** 3- or 4-card row. Each card: picture (6×6cm); name (20pt bold); role (16pt); **prior company + title** (italic); optional LinkedIn footer.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[11]",
      "--type",
      "shape",
      "--prop",
      "text=Team: 3 prior exits, 42 years combined K8s",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[11]",
      "--type",
      "picture",
      "--prop",
      "src=/workspace/assets/alice.jpg",
      "--prop",
      "x=2cm",
      "--prop",
      "y=5cm",
      "--prop",
      "width=6cm",
      "--prop",
      "height=6cm"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[11]/picture[1]",
      "--prop",
      "alt=Alice Chen, CEO — portrait"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[11]",
      "--type",
      "shape",
      "--prop",
      "text=Alice Chen",
      "--prop",
      "x=2cm",
      "--prop",
      "y=11.5cm",
      "--prop",
      "width=6cm",
      "--prop",
      "height=1cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=20",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[11]",
      "--type",
      "shape",
      "--prop",
      "text=CEO",
      "--prop",
      "x=2cm",
      "--prop",
      "y=12.8cm",
      "--prop",
      "width=6cm",
      "--prop",
      "height=0.8cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=16",
      "--prop",
      "color=333333",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[11]",
      "--type",
      "shape",
      "--prop",
      "text=ex-Datadog Director (Series C → IPO); led K8s observability GTM $40M → $200M ARR",
      "--prop",
      "x=2cm",
      "--prop",
      "y=13.8cm",
      "--prop",
      "width=6cm",
      "--prop",
      "height=2.5cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=14",
      "--prop",
      "italic=true",
      "--prop",
      "color=333333",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

Repeat for Card 2 (CTO, x=`10cm`) and Card 3 (VP Eng, x=`18cm`) — 3 cards × 5–6 shapes each.

**Arrangement helper.** 3 cards: `col=9.78cm, x=1.5/12.04/22.58`. 4 cards: `col=7.15cm, x=1.5/9.41/17.32/25.23`. 5+: see pptx grid math.

**QA.** Query `shape:contains("ex-")` / `"prior"` / `"former"` ≥ 1 per team member.

### (9) Financials slide — 4-year plan + honest assumptions

**Visual outcome.** Column chart: 4 years × (revenue, GM $, EBITDA). Right-side assumptions panel. Title names the trajectory.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[17]",
      "--type",
      "shape",
      "--prop",
      "text=$18M → $85M ARR by FY29",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[17]",
      "--type",
      "chart",
      "--prop",
      "chartType=column",
      "--prop",
      "series1.name=Revenue ($M)",
      "--prop",
      "series1.values=18,34,58,85",
      "--prop",
      "series1.color=1E2761",
      "--prop",
      "series2.name=Gross Margin ($M)",
      "--prop",
      "series2.values=14,26,45,68",
      "--prop",
      "series2.color=CADCFC",
      "--prop",
      "series3.name=EBITDA ($M)",
      "--prop",
      "series3.values=-6,-2,8,22",
      "--prop",
      "series3.color=B85042",
      "--prop",
      "categories=FY26,FY27,FY28,FY29",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=20cm",
      "--prop",
      "height=13cm",
      "--prop",
      "title=4-year plan — revenue, GM, EBITDA ($M)"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[17]",
      "--type",
      "shape",
      "--prop",
      "geometry=roundRect",
      "--prop",
      "fill=F5F7FA",
      "--prop",
      "line=none",
      "--prop",
      "x=22.5cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=9.8cm",
      "--prop",
      "height=13cm"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[17]",
      "--type",
      "shape",
      "--prop",
      "text=Key Assumptions",
      "--prop",
      "x=23cm",
      "--prop",
      "y=4.5cm",
      "--prop",
      "width=8.8cm",
      "--prop",
      "height=1.2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=20",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

Add 5 assumption bullets as separate paragraph shapes at y=`6, 7.5, 9, 10.5, 12cm` — size=14, italic=true; each ≤ 14 words.

**QA.** Query `shape:contains("assumption")` OR `contains("Assumes")` ≥ 1.

### (10) The Ask — hero number + 4-bucket Use-of-Funds + runway

**Visual outcome.** Dark fill (match cover). Hero `$35M` ~88–96pt white. Below: 4-bucket pie (Engineering 40% / GTM 35% / G&A 15% / Reserve 10%). Bottom: runway + next milestone.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=1E2761"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[20]",
      "--type",
      "shape",
      "--prop",
      "text=$35M Series B",
      "--prop",
      "x=2cm",
      "--prop",
      "y=2cm",
      "--prop",
      "width=29.87cm",
      "--prop",
      "height=4cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=88",
      "--prop",
      "bold=true",
      "--prop",
      "color=FFFFFF",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[20]",
      "--type",
      "chart",
      "--prop",
      "chartType=pie",
      "--prop",
      "series1.name=Use of Funds",
      "--prop",
      "series1.values=40,35,15,10",
      "--prop",
      "categories=Engineering,Go-to-Market,G&A,Reserve",
      "--prop",
      "colors=CADCFC,B85042,97BC62,FFFFFF",
      "--prop",
      "x=6cm",
      "--prop",
      "y=7cm",
      "--prop",
      "width=12cm",
      "--prop",
      "height=10cm",
      "--prop",
      "title=Use of Funds — 4 buckets"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[20]",
      "--type",
      "shape",
      "--prop",
      "text=18 months runway to $40M ARR and Series C",
      "--prop",
      "x=2cm",
      "--prop",
      "y=17cm",
      "--prop",
      "width=29.87cm",
      "--prop",
      "height=1.5cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=22",
      "--prop",
      "color=CADCFC",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
    ]
  }
}
```

**4-bucket convention.** Engineering / GTM / G&A / Reserve. Typical Series A: Eng 40-50%, GTM 30-40%, G&A 10-15%, Reserve 5-10%. Series B shifts 5-10 points from Eng to GTM.

**QA.** Query `shape:contains("Use of Funds")` ≥ 1. Pie present. Runway + milestone on ask slide.

### (11) Pipeline chart — Bio / Deep Tech must-have

**Visual outcome.** Horizontal swimlane. Left = candidate; 4 stage columns (Preclinical / Ph1 / Ph2 / Ph3 or TRL bands). NCT footer. SaaS / Consumer skip.

Grid: usable `= 30.87cm`, candidate col `= 7cm`, stage cols `= 5.97cm`, row height `= 2.3cm`. Stage x: `8.5 / 14.47 / 20.44 / 26.41`.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "shape",
      "--prop",
      "text=Pipeline: 3 candidates across Ph1–Ph3",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "shape",
      "--prop",
      "text=Preclinical",
      "--prop",
      "x=8.5cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=5.97cm",
      "--prop",
      "height=1cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=16",
      "--prop",
      "bold=true",
      "--prop",
      "color=333333",
      "--prop",
      "align=center",
      "--prop",
      "fill=F5F7FA"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "shape",
      "--prop",
      "text=Phase 1",
      "--prop",
      "x=14.47cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=5.97cm",
      "--prop",
      "height=1cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=16",
      "--prop",
      "bold=true",
      "--prop",
      "color=333333",
      "--prop",
      "align=center",
      "--prop",
      "fill=F5F7FA"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "shape",
      "--prop",
      "text=Phase 2",
      "--prop",
      "x=20.44cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=5.97cm",
      "--prop",
      "height=1cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=16",
      "--prop",
      "bold=true",
      "--prop",
      "color=333333",
      "--prop",
      "align=center",
      "--prop",
      "fill=F5F7FA"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "shape",
      "--prop",
      "text=Phase 3",
      "--prop",
      "x=26.41cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=5.97cm",
      "--prop",
      "height=1cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=16",
      "--prop",
      "bold=true",
      "--prop",
      "color=333333",
      "--prop",
      "align=center",
      "--prop",
      "fill=F5F7FA"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "shape",
      "--prop",
      "text=HLX-201 (lead)",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=5.5cm",
      "--prop",
      "width=7cm",
      "--prop",
      "height=1.5cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=18",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "align=left",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[6]",
      "--type",
      "shape",
      "--prop",
      "geometry=roundRect",
      "--prop",
      "fill=1E2761",
      "--prop",
      "x=8.5cm",
      "--prop",
      "y=5.7cm",
      "--prop",
      "width=17.91cm",
      "--prop",
      "height=1.1cm",
      "--prop",
      "line=none"
    ]
  }
}
```

Repeat rows 2 & 3 at y=`7.8cm` / y=`10.1cm` with bar widths per stage (Ph1=`5.97cm`, Ph1-Ph2=`11.94cm`, Ph1-Ph3=`17.91cm`). NCT footer full-width at y=`16.8cm` with trial IDs.

**QA.** Query `shape:contains("NCT")` ≥ 1. Bar colors darken across stages.

### (12) Competitive comparison table — Series B+ essential

**Visual outcome.** 5–7 rows × 4–6 cols. Last row = your company, highlighted. Every Series B+ deck needs this.

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/",
      "--type",
      "slide",
      "--prop",
      "layout=blank",
      "--prop",
      "background=FFFFFF"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[13]",
      "--type",
      "shape",
      "--prop",
      "text=Competitive landscape",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761",
      "--prop",
      "fill=none"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "/slide[13]",
      "--type",
      "table",
      "--prop",
      "data=Competitor,Speed,Price,Integrations,Margin;Datadog,12 min,$15/host,680,75%;New Relic,18 min,$25/host,520,68%;Splunk,45 min,$45/GB,310,62%;You (Acme DevOps),90 sec,$8/host,1200,82%",
      "--prop",
      "style=medium1",
      "--prop",
      "headerFill=1E2761",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=30.87cm",
      "--prop",
      "height=12cm"
    ]
  }
}
```

Highlight your row: set cell fill on `/slide[13]/table[1]/tr[5]/tc[1..5]` to `CADCFC` (separate Run/Batch after table exists).

**QA.** Query `table` length ≥ 1. Row count ≥ 4. Your row visually distinct (Gate 5b).

## Numbers convention (pitch-specific)

A terse convention table — **not a finance tutorial**. If you don't already know what these mean, pause and ask the user; don't guess.

| Metric | Shape | Floor / convention |
|---|---|---|
| **TAM** | `$X.YB`, one methodology | Either top-down or bottom-up. Never both; never neither. |
| **SAM** | `$X.YB`, fraction of TAM you serve | Typically 15 – 30% of TAM for verticalized SaaS |
| **SOM** | `$X.YB` at year N | Realistic 5-yr share: 5 – 15% of SAM for early stage |
| **ARR** | MRR × 12. NOT revenue. | SaaS only; contracts on books, net of churn |
| **MRR** | Monthly recurring | ARR / 12 |
| **NRR** | %, trailing 12 mo | > 100% acceptable, > 115% strong, > 130% exceptional |
| **CAC** | $ fully-loaded | S&M spend / new logos |
| **LTV** | $ | ARPU × GM × (1 / churn) |
| **LTV:CAC** | ratio | 3x OK, > 4x strong, > 5x exceptional |
| **CAC payback** | months | < 18 mo OK, < 12 mo strong |
| **Gross margin** | % | SaaS floor 70%; marketplace 15-40%; hardware 30-50% |
| **Burn / runway** | $/month + months | Label gross vs net; runway to specific milestone |
| **Use of Funds** | 4-bucket pie | Eng / GTM / G&A / Reserve |

**Rule.** Every number carries a unit. `TBD`, `coming soon`, `(fill in)`, `lorem`, `xxxx` in numeric slots = immediate VC disqualification.

## VC ship-check (6 red flags / positive signals)

| # | Red flag (FAIL if present) | Positive signal |
|---|---|---|
| 1 | Cover without round + amount + date | `Company · tagline · Series X · $YM · Date` |
| 2 | TAM > $100B without cited source / methodology | TAM labeled bottom-up OR top-down with 2024+ source |
| 3 | Traction chart y-axis does not start at 0 | Line chart `axismin=0` |
| 4 | Team: headshots + names only | Every member: prior company + role + 1 achievement |
| 5 | Ask missing Use-of-Funds | `$XM` hero + 4-bucket + runway + next milestone |
| 6 | `TBD` / `lorem` / `xxxx` / `{{...}}` / `(fill in)` | `view text` clean |

**Series-specific failures.** A: fictional bottom-up / premature CAC. B: no unit-econ / payback >24mo / logo wall <8. C: no moat / growth without margin.

## Traction triple-pattern (ARR + milestones + logos)

Series B+ often needs a second traction slide: **milestone timeline + logo wall**. Timeline = 4-6 dates (→ pptx Roadmap timeline). Logo wall = 12-20 muted logos; 5-across grid math: `usable = 33.87 − 3 − 4·0.4 = 29.27`, `col = 5.85cm`. **QA:** ≥ 8 logos Series B+, ≥ 4 Series A; >20 = noise.

## QA — Delivery Gate (executable)

**Assume there are problems.** First render is almost never correct. Pitch decks fail at structural (Gates 1–3) and narrative (Gate 5b + Gate 6) layers.

### Gates 1–5a — inherited from pptx verbatim

→ see pptx §Delivery Gate. Gate 1 `validate` (whitelist ChartShapeProperties per C-P-2). Gate 2 token leak via `view text`. Gate 3 hyperlink rPr. Gate 4 slide-order. Gate 5a dark-on-dark (including chart title/legend/axis on dark fills).

**Gate 2b — pitch-specific $-strip signatures (MANDATORY).** Gate 2 misses `$35M` silently stripped. After Gate 2, Run `view text` and reject patterns like bare `M ARR` / `Series [A-C] · M` / `runway · M`. Fix by re-issuing the offending `add`/`set` with `$` preserved in JSON strings. Same strip hits chart series names / axis titles carrying `$`.

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "text"
    ]
  }
}
```

### Gate 5b — Visual audit via HTML preview (MANDATORY)

Run `view html` and Read the returned HTML. For EACH slide (inherits pptx Gate 5b; pitch additions ⭐):

- overlap / dark-on-dark / divider overlap / order sanity / missing arrowheads
- ⭐ traction y-axis starts at 0
- ⭐ team credibility (prior company / title per card)
- ⭐ TAM credibility (under $100B OR methodology cited)
- ⭐ Use-of-Funds pie / 4-card %s on ask
- ⭐ narrative completeness (cover → … → ask, or stage-appropriate permutation)

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "html"
    ]
  }
}
```

Report every defect with slide number. If ANY defect — REJECT; do not deliver until fixed.

### Gate 6 — Pitch narrative sanity

Token checks — combine with Gate 5b:

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": [
      "/workspace/pitch.pptx"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": [
      "/workspace/pitch.pptx",
      "shape:contains(\"Use of Funds\")",
      "--json"
    ]
  }
}
```

Checks:
- **6.1** no TBD / lorem / (fill in) / xxxx / coming soon / placeholder in `view text`
- **6.2** TAM presence (Series A+) — WARN if intentional Seed/Bridge
- **6.3** CAC OR LTV (Series B+) — WARN/REJECT by stage; non-SaaS use vertical substitutes
- **6.4** Use of Funds ≥ 1 — REJECT if missing
- **6.5** team prior-company signal (`ex-` / former / prior / previously) — REJECT if zero
- **6.6** at least one chart with `axisMin`/`axismin` = 0 (Series A+) — WARN if no ARR/revenue line chart

**Readback key note.** Input prop is lowercase `axismin`; readback often camelCase `axisMin`. Accept both.

Gate 6 is a grep floor. Gate 5b is the visual ceiling. Ship only when both PASS.

### Honest limit

`validate` catches schema errors, not fundraising errors. Gates 5b + 6 exist because `validate` cannot catch TAM fabrication, missing prior companies, hockey-stick axes, missing unit econ, or a vague ask.

## Known Issues & Pitfalls

→ Base pitfalls: see pptx §Known Issues C-P-1..7.

Pitch-specific:

- **Stage misidentified.** Series A with 6 pages of CAC/LTV = over-packaged. Series B missing unit econ = incomplete.
- **Hockey-stick y-axis.** Always `axismin=0` on ARR / revenue / growth charts. Gate 6.6.
- **Team slide = portfolio.** Every card needs prior-company or prior-achievement. Gate 6.5.
- **TAM without methodology.** Pick one methodology per deck; don't mix.
- **Use-of-Funds as 3-bucket or 5-bucket.** 4-bucket is convention. Gate 6.4.
- **Pitch deck used for board review / sales.** Route to pptx. See §Reverse handoff.
- **pptx 20-slide blueprint is a starting point, not a formula.** Adjust for stage + 赛道 — never ship unchanged for non-SaaS Series A.

## Help pointer

When in doubt: Help pptx / pptx `<element>` / `json` via `{{OFFICE_HELP_TOOL}}`. Help is the authoritative schema; this skill is the decision guide for fundraising deltas on top of pptx.
