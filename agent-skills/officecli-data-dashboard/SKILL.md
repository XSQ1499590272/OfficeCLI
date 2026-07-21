# Data Dashboard Agent Skill

Author: xiesq, 2026-07-15

**Do not invent** metrics, KPI values, chart series, or board-pack numbers. Bind only to user-supplied or formula-derived cells.

Dashboard ≠ “带 chart 的表”。它是组合：**打开即见的 Dashboard sheet** + 公式驱动 KPI cards + cell-range 链接 charts + sparklines + 语义 CF。原始数据与聚合是上游基础设施。xlsx 引擎规则全部继承，不重教。

## ⚠️ Help 优先规则

不确定的 prop / enum / alias，先查 Help。

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "xlsx"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "xlsx",
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
      "xlsx",
      "sparkline"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "xlsx",
      "conditionalformatting"
    ]
  }
}
```

Skill 与 Help 冲突时，**以 Help 为准**。DeferredAddKeys（`combosplit`、`holesize`）仅 `add` 有效——见 Known Issues。

```json
{
  "tool": "{{OFFICE_LOAD_SKILL_TOOL}}",
  "arguments": {
    "command_arguments": [
      "excel"
    ]
  }
}
```

`{{OFFICE_LOAD_SKILL_TOOL}}` 只读取指导。先读 `excel` / excel 基座。

## 心智模型与继承

继承 xlsx 硬规则：零 formula error、视觉底线、Batch dotted-name、chart data-feed、validate、HTML preview。不重教。

### Reverse handoff

- 单 sheet CSV tracker、≤1 chart → `excel`/`xlsx` 基座
- 3-statement / DCF / LBO → `financial-model`
- weekly status 一行 SUMIF + 一图 <10 行 → 基座

仅当：Dashboard 首页 + 多 KPI + 多 chart + 若干 CF/sparkline。

## 宿主工具与执行规范

（对应原 CLI「Shell & Execution Discipline」；shell 引号改为 Tool JSON 传参。）

**Path 与特殊字符。** Excel path 含 `[]`，number format 可能含 `$`，formula 可能含跨 sheet `!`：path / props 原样写入 JSON 字符串即可。

**增量执行。** 长 chart `add` 用独立 Run/Batch 项，勿把整条塞进难诊断的超大块。多实例计数用 `query json` 读 `.data.results | length`，勿 `raw` + grep。节奏：一条（或一块独立 Batch）→ 检查 → 继续。生命周期 `create` / `open` / `save` / `close`。

**阶段 0 Batch/Run 契约（固定补丁，非业务加戏）。**

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。

## Core Principles（五条不可协商）

1. **Formula-driven KPIs。** Dashboard 上每个 KPI 值是公式引用 Data/Summary，永不硬编码计算结果。
2. **Cell-range charts。** `series1.values="Sheet1!B2:B13"`。Inline `data=` 只给 5 分钟 demo。
3. **Dashboard-first。** KPI/label/chart/sparkline 都在 Dashboard；用户不应为答案切 tab。
4. **Visible cells only for chart sources。** LibreOffice 不对隐藏列/表在渲染时求值——聚合到可见 Summary，再指向 Summary。
5. **Data-size-aware complexity。** 10 行不要 5 KPI+4 charts；200 行不要 1 KPI+1 chart。

## Requirements

继承 xlsx 要求，另加：

- Dashboard 是打开时 active tab（用 `query sheet` 确认 0-based index，再设 activeTab——不要猜）。
- `calc.fullCalcOnLoad=true` 用高层 `set`，不要 `raw-set` 造重复 `<calcPr>`。
- 上游改完后刷新下游 formula/cachedValue（fullCalcOnLoad 只管 runtime）。
- 每个 chart 有描述性 title；每个 series 有 name（禁 “Series1”）。
- 每个 KPI value 有 formula（可 query `Dashboard!:has(formula)` 计数）。
- Data/Summary header fill（如 `1F3864` + 白字 bold）。
- 10+ 行 Data → ≥1 CF on numeric column。
- KPI 列宽按最大 cachedValue 估算：`width ≈ ceil((visible_chars+2)*1.3)`；货币长数字常需 28–44，不要死写 22。
- Sparkline 行高 ≥20（常 22–24）。
- Print 交付：`_xlnm.Print_Area` 限定 Dashboard + 隐藏非 Dashboard + fitToPage（需 Help 确认的高层 props；若必须 raw XML → prose + **approval** + 逐步 Run，**永不**进 Batch JSON）。

## Quick Start

最小可行：12 月 revenue CSV → 4 KPI 中先做 1 个 + 1 line chart + fullCalcOnLoad + activeTab。按实际数据调整，勿盲贴。

### Phase 1 — Data sheet: create, import, format

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": ["/workspace/dashboard.xlsx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "import",
    "command_arguments": [
      "/workspace/dashboard.xlsx",
      "/Sheet1",
      "--file", "sales.csv",
      "--header"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/dashboard.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/Sheet1/col[A]",
        "props": { "width": 12 }
      },
      {
        "command": "set",
        "path": "/Sheet1/col[B]",
        "props": { "width": 15 }
      },
      {
        "command": "set",
        "path": "/Sheet1/B2:B13",
        "props": { "numFmt": "$#,##0" }
      },
      {
        "command": "set",
        "path": "/Sheet1/A1:B1",
        "props": {
          "fill": "1F3864",
          "font.color": "FFFFFF",
          "font.bold": true
        }
      }
    ],
    "stop_on_error": true
  }
}
```

`import header` 设 freeze + AutoFilter，但不设 column width / 完整 `numFmt`（见 D-12）。

### Phase 2 — Dashboard sheet + one KPI card

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/dashboard.xlsx",
      "/",
      "--type",
      "sheet",
      "--prop",
      "name=Dashboard"
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
      "/workspace/dashboard.xlsx",
      "/Dashboard/col[A]",
      "--prop",
      "width=22"
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
      "/workspace/dashboard.xlsx",
      "/Dashboard/col[B]",
      "--prop",
      "width=12"
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
      "/workspace/dashboard.xlsx",
      "/Dashboard/A1",
      "--prop",
      "value=Total Revenue",
      "--prop",
      "font.size=9",
      "--prop",
      "font.color=666666",
      "--prop",
      "bold=true"
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
      "/workspace/dashboard.xlsx",
      "/Dashboard/A2",
      "--prop",
      "formula=SUM(Sheet1!B2:B13)",
      "--prop",
      "numFmt=$#,##0",
      "--prop",
      "font.size=24",
      "--prop",
      "bold=true",
      "--prop",
      "font.color=2E7D32"
    ]
  }
}
```

### Phase 3 — Sparkline + chart + fullCalcOnLoad

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/dashboard.xlsx",
    "operations": [
      {
        "command": "add",
        "parent": "/Dashboard",
        "type": "sparkline",
        "props": {
          "cell": "B2",
          "range": "Sheet1!B2:B13",
          "type": "line",
          "color": "4472C4",
          "highPoint": true,
          "highMarkerColor": "FF0000"
        }
      },
      {
        "command": "add",
        "parent": "/Dashboard",
        "type": "chart",
        "props": {
          "chartType": "line",
          "title": "Revenue Trend",
          "series1.name": "Revenue",
          "series1.values": "Sheet1!B2:B13",
          "series1.categories": "Sheet1!A2:A13",
          "preset": "dashboard",
          "axisNumFmt": "$#,##0",
          "x": 0,
          "y": 5,
          "width": 10,
          "height": 15
        }
      },
      {
        "command": "set",
        "path": "/",
        "props": { "calc.fullCalcOnLoad": true }
      }
    ],
    "stop_on_error": true
  }
}
```

### Phase 4 — activeTab LAST → close → validate

先 Run `query sheet` 解析 Dashboard 的 0-based index（勿硬编码）。再用 Help 确认的方式设置 `activeTab`（若需 workbookView raw 插入 → prose + **approval** + 逐步 Run，依赖 **relationship**/path，**不进 Batch JSON**）。然后：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": ["/workspace/dashboard.xlsx", "/Dashboard/A2"]
  }
}
```

确认 `A2` 有 formula 与合理 `cachedValue`（测试数据示例约 2,075,000）。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "close",
    "command_arguments": ["/workspace/dashboard.xlsx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": ["/workspace/dashboard.xlsx"]
  }
}
```

## Design Ideas

### Layout patterns（选一，保持一致）

**1 Executive summary（board）：** KPI strip A1:H4；charts from row 6（宽图 + 下排双图）。
**2 Ops console：** KPIs 在 A:B；charts 填 C:L。
**3 Scorecard（≥6 KPI）：** 2×3 卡（label / value / sparkline）。

### Complexity scaling by data size

| Rows | KPIs | Charts | Sparklines | CF | Preset |
|---|---|---|---|---|---|
| <10 | 1–2 | 1 | skip | 0–1 | minimal |
| 10–50 | 2–3 | 2 | 仅有序时间序列 | 1–2 | dashboard |
| 50–200 | 3–5 | 2–3 | 同上 | 2–3 | dashboard |
| 200+ | 3–5 | 3 | 同上 | 3–4 | dashboard |

### Chart type selection

| Pattern | Type | Notes |
|---|---|---|
| 单序列趋势 | line | 可 `trendline=linear` |
| 多组件趋势 | line multi / columnStacked | 组件可加总时用 stacked |
| 类别比较（时间序） | column | 非 bar（破坏时间阅读） |
| 部分-整体 | doughnut | `pie` 在 LibreOffice 有 blank 回归 |
| Budget vs actual | combo + `combosplit=1` | DeferredAddKey：仅 add 时 |
| Correlation | scatter | x 轴用 `categories`；`series1.xValues` UNSUPPORTED |

Preset 全 Dashboard 一致：`minimal` / `dashboard` / `corporate` / `magazine` / `colorful` / `monochrome` / `dark`。

### Conditional formatting — semantic colors

| Intent | type | Typical props |
|---|---|---|
| Magnitude | databar | `sqref` + `color`；可选 min/max |
| Heat map | colorscale | min/mid/max colors |
| Status | iconset | `3Arrows` 等（见 Help） |
| Business rule | formulacf | `formula` + fill/font（`font.bold` 可用） |

语义色：好 `C8E6C9`/`2E7D32`；坏 `FFCDD2`/`C62828`；中性 `F5F5F5`/`666666`。

10+ 行 Data 时加 CF（示例）：

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/dashboard.xlsx",
    "operations": [
      {
        "command": "add",
        "parent": "/Sheet1",
        "type": "conditionalformatting",
        "props": { "type": "databar", "sqref": "B2:B13", "color": "4472C4" }
      },
      {
        "command": "set",
        "path": "/Dashboard/row[2]",
        "props": { "height": 22 }
      },
      {
        "command": "set",
        "path": "/Dashboard/col[A]",
        "props": { "width": 28 }
      }
    ],
    "stop_on_error": true
  }
}
```

### KPI card anatomy

Label：size 9 gray bold。Value：size 24 bold + numFmt + tone color。可选浅底 `F0F4FF`。列宽跟最大 cachedValue。

### Chart width vs title length

经验：`chart.width ≥ ceil(title.length × 0.18)`。太长则加宽或缩短标题 ≤25 字。`get chart[N]` 可读 `width` 与 `anchor`。

### Print-ready delivery

触发：print / 一页 / 董事会 / 投资人。四件套：Print_Area 限定 Dashboard；fit-to-page；landscape；隐藏非 Dashboard（Print_Area alone 不阻止其它 sheet 进打印管道）。删掉 Data/Summary 上冲突的 Print_Area。

高层 props 优先；pageSetup / sheet hidden 若必须 raw XML → prose + **approval** + 逐步 Run（**relationship**/path），**永不**进 Batch JSON。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/dashboard.xlsx",
    "operations": [
      {
        "command": "add",
        "parent": "/",
        "type": "namedrange",
        "props": {
          "name": "_xlnm.Print_Area",
          "scope": "Dashboard",
          "refersTo": "Dashboard!$A$1:$H$36"
        }
      },
      {
        "command": "set",
        "path": "/Dashboard",
        "props": { "orientation": "landscape", "fitToPage": true }
      },
      {
        "command": "set",
        "path": "/Dashboard/col[A]",
        "props": { "width": 28 }
      }
    ],
    "stop_on_error": true
  }
}
```

`orientation` / `fitToPage` 以 Help `xlsx` sheet props 为准；若宿主版本无高层 alias，改用 prose 描述的 raw 路径（仍需 **approval**）。

## QA（REQUIRED — Delivery Gate）

**假定存在问题。** validate 通过 ≠ 可交付。

继承 xlsx minimum cycle（`view issues`、error query、`validate`、HTML），再跑：

**Gate 1** KPI formula 覆盖（计划 N 个则 ≥N）：`query Dashboard!:has(formula)`。
**Gate 2** chart 数量、有 data、title 非空、anchor span 够 title（约 title.length×0.18）。
**Gate 3** 无 auto-named `SeriesN`。
**Gate 4** 10+ 行时 CF ≥1（element 名 `conditionalformatting`，`query cf` 不是 alias）。
**Gate 5** activeTab == Dashboard index；`calc.fullCalcOnLoad` true。
**Gate 6** placeholder sweep（`view text`：`{{` / `$fy$` / `<TODO>` / `xxxx` / `TBD`）。
**Gate 7** `view html`：无 `###`、无裁切、无空 chart、Dashboard 打开第一；print 场景检查非 Dashboard 已隐藏 + Print_Area。
**Gate 8** 每个 KPI formula 的 cachedValue 非空/非错误/默认非零（真零需白名单说明）。`fullCalcOnLoad` 不刷 build-time cache。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": ["/workspace/dashboard.xlsx", "html"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": [
      "/workspace/dashboard.xlsx",
      "Dashboard!:has(formula)",
      "--json"
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
      "/workspace/dashboard.xlsx",
      "conditionalformatting",
      "--json"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": [
      "/workspace/dashboard.xlsx",
      "/Dashboard/chart[1]",
      "--json"
    ]
  }
}
```

Gate 2/3：读 `seriesCount`、`title`、`anchor`、series `name`（拒 `SeriesN`）。Gate 5：`query sheet` 得 Dashboard 索引后与 workbook `activeTab` / `calc.fullCalcOnLoad` 比对。失败则源头修复并重跑全循环。

### Honest limits

Scatter 无 `xValues`；LibreOffice 颜色漂移 / pie 塌陷是 viewer artifact——先在目标 viewer 确认。

## Reference

- Shorthand `type` at add：`chart` / `sparkline` / `databar` / `colorscale` / `iconset` / `formulacf`。CF path `/Sheet/cf[N]`。
- Full schema：Help `chart` / `sparkline` / `conditionalformatting`。
- DeferredAddKeys（add-only）：`combosplit`、`holesize`。`preset` / `trendline` / `referenceline` / `axisNumFmt` 现已支持 set（Help 标 `[add/set]`）。
- Build order：charts + sparklines + CF + tabColors → `calc.fullCalcOnLoad=true` → `activeTab` **LAST**。

## Known Issues（D-1..D-17 摘要）

| # | Issue | Mitigation |
|---|---|---|
| D-1 | `combosplit` 仅 add | add 时设置；其余四键可 set |
| D-2 | `referenceline` 顺序 value:color:label:dash | 勿颠倒 |
| D-3 | scatter 无 `xValues` | `categories` / `series1.categories` |
| D-4 | `formulacf` 支持 font.bold/italic | 与 fill/font.color 同用 |
| D-5 | 默认列宽 8.43 → `###` | 按 cachedValue 加宽（亿级常 32–44） |
| D-6 | activeTab 必须最后 | 全部 sheet/chart 后再设 |
| D-7 | raw `<calcPr>` 重复 | 用高层 `set calc.fullCalcOnLoad` |
| D-8 | 隐藏源 → 空白 chart | 可见 Summary |
| D-9 | `pie` blank LO | 用 doughnut |
| D-10 | SUMIFS 日期准则字符串静默失败 | `DATE()` / `DATEVALUE()` |
| D-11 | 百分比无 numFmt 显示小数 | 同次 set `numFmt=0.0%` |
| D-12 | import header 不设 width | 另 set `col[]` width；列路径的 `numFmt` 只提供列级 style 并覆盖空白 cell，已有自有 style 的 cell 仍需对目标 cell range 单独设置 `numFmt` |
| D-13 | `highPoint` 是 bool | `highPoint=true` + `highMarkerColor` |
| D-14 | 无序类别不要 sparkline | 仅时间序列 |
| D-15 | 空 chart add 被拒 | 必须 series/dataRange/data |
| D-16 | fullCalcOnLoad 不刷 build cache | 重设公式或 close/reopen；Gate 8 |
| D-17 | SUMPRODUCT 数组谓词可能 cache 0 | helper 列 + SUMIF |

跨表 `!`、chart series 不可变 → xlsx Known Issues。超出 Help 的 workbook XML：`raw-set`/`add-part` 仅 prose + **approval**；**relationship**/path 依赖逐步 Run；**永不**进 Batch JSON。
