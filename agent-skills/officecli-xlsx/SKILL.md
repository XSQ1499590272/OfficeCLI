# XLSX Agent Skill

Author: xiesq, 2026-07-15

## ⚠️ Help 优先规则

**本 Skill 说明高质量 xlsx 应达到的标准，而不是罗列每一个 command flag。遇到不确定的 property name、enum value 或 alias 时，先查 Help，切勿猜测。**

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": ["xlsx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": ["xlsx", "chart"]
  }
}
```

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": ["xlsx", "add", "chart"]
  }
}
```

Help 与已安装的 Office 工具版本一致。如本 Skill 与 Help 不一致，**以 Help 为准**。

`{{OFFICE_LOAD_SKILL_TOOL}}` 只读取指导，不安装或修改本地 skills。

## 宿主工具与执行规范

（对应原 CLI「Shell 与执行规范」；shell 引号改为 Tool JSON 传参。）

**Path 与特殊字符。** Excel path 含 `[]`，number format 可能含 `$`，formula 可能含跨 sheet `!`：

- path 原样传入 `command_arguments` / Batch `path`（含 `[N]`），无需 shell 引号。
- `$`、`!` 写在 JSON 字符串内即可（无 shell history expansion）。
- props 中的 `\n` / `\t` 仍由 Office 工具解释：`\n` 为 cell 内换行（配合 `wrapText=true`），`\t` 为 tab；字面量反斜杠+n 写 `\\n`。

**增量执行。** 每次结构操作后检查结果再继续。含 50 条操作的流程若第 3 条失败，后续会连锁失败。正确节奏：一条（或一块独立 Batch）→ 检查 output → 继续。

**阶段 0 Batch/Run 契约（固定补丁，非业务加戏）。**

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。

## 输出标准

执行前，先理解高质量 xlsx 的定义。以下是每个 workbook 必须达到的交付标准。

### 所有 Excel 文件

**formula error 必须为零。** 每个交付的 workbook 必须没有 `#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#N/A`。没有例外；使用 `IFERROR` 或 `IF(x=0,...)` 保护分母。

**应使用 formula，而非 hardcoded value。** 数字若可由其他 cell 计算得出，就应写为 formula。在应使用 `SUM(B2:B9)` 的位置硬编码 `5000`，会破坏 workbook 随输入变化而保持 live 的约定。这是本 Skill 最重要的一条规则。

**专业 font。** workbook 中应统一使用一种专业 font（Arial / Calibri / Times New Roman）。不要因为某个 sheet 来自 CSV 就混用四种 font。

**显式 column width。** 不存在 auto-fit。用户会阅读的 column 必须设置 `width`；默认 8.43 个字符会截断内容。合理起点：label 20–25、number 12–15、date 12、short code 8–10。

**保留现有 template。** 编辑已有视觉样式的文件时应遵循原有约定；它们优先于本指南。

### 视觉交付底线（适用于每个 workbook）

声明完成前，Run `view html` 并读取返回的 HTML path，确认以下全部条件：

- **任何 cell 中不得出现 `###`。** `###` 表示 column 宽度不足以展示最长 value。用户会阅读的每个 column 都需要显式 `width`。交付文件中的 `###` 是未完成工作，不是“小的视觉瑕疵”。
- **不得截断 title。** sheet title、section header、long label 都必须完整显示；加宽 column 或在 cell 上设置 `wrapText=true`。
- **不得将 placeholder token 渲染为数据。** `$fy$24`、`{var}`、`<TODO>`、`xxxx` 不得出现在 cell、chart title、series name 或 legend 中；它们是未被替换的 build-time token。
- **pie / doughnut slice 应使用不同 fill color。** 若 slice 渲染成同色，切换为 `bar` / `column`，或显式设置 `colors=...`。
- **不得存在空尾页或空 chart anchor。** 例如，源 cell 为空时的 `anchor=D2:J18` 看起来像损坏的 chart。

如果上述任何一项失败，请在宣布完成之前停止并修复。

**Print layout。** 用户可能打印或作为 board pack 发送的任何 sheet 都需要 page setup。默认 portrait 且不 fit-to-page 会把宽 table 和 chart 拆到多页。应按 sheet 形状选择 fit mode：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/book.xlsx",
      "/Summary",
      "--prop", "orientation=landscape",
      "--prop", "fitToPage=true"
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
      "/workspace/book.xlsx",
      "/Data",
      "--prop", "orientation=landscape",
      "--prop", "fitToPage=1x0"
    ]
  }
}
```

`fitToPage=true` 等于 `1x1`，即宽高都压到一页，只适合原本就短的 sheet；`1x0` 表示宽度压到 1 页、高度不限。触发条件：sheet 含 chart、超过 8 个 column，或用户需求提到 print / board / investor。

### 仅限 financial model：template、tracker、CSV import 或 operational sheet 可跳过

适用范围：budget、forecast、three-statement model、valuation，以及任何 `$` 密集的 analytical workbook。customer-support tracker 或 onboarding template 无需遵循本节。

**Color coding：行业标准。** 五种核心颜色是信息语言，不是装饰。reviewer 应在阅读 formula 前，仅凭颜色就能判断 cell 的类型。

| Color | 角色 | 示例 |
|---|---|---|
| Blue text `0000FF` | hardcoded input、scenario variable | `font.color=0000FF` |
| Black text `000000` | 所有 formula 与 calculation | 默认 |
| Green text `008000` | 当前 workbook 内的 cross-sheet link | `font.color=008000` |
| Red text `FF0000` | 指向 external file / workbook 的 link | `font.color=FF0000` |
| Yellow fill `FFFF00` | 需要 review 的 key assumption | `fill=FFFF00` |

**Number format：标准，而非偏好。**

- **年份**应为 text 而非 number：显示 `2026`，不要显示 `2,026`；使用 `numFmt="@"` 或设置 `type=string`。
- **货币**单位写在 header 中（`Revenue ($mm)`），不要在每个 cell 中重复。
- **零显示为 `-`**，而不是 `0`：使用 `$#,##0;($#,##0);"-"`。
- **百分比**默认保留一位小数：`0.0%`。
- **负数用括号**：使用 `(1,234)`，不要使用 `-1,234`。
- **Valuation multiple** 使用 `0.0x` format（EV/EBITDA、P/E 等）。

**Assumption 应存放在 cell 中，不要硬编码在 formula 内。** `=B5*(1+$B$6)` 正确；`=B5*1.05` 是 bug。每个 blue hardcoded input 均应在相邻 cell 或 cell comment 中记录 source note：

```
Source: Company 10-K, FY2024, Page 45, Revenue Note
Source: Bloomberg, 2026-05-02, AAPL US Equity
Source: Management guidance, Q2 2026 earnings call
```

任何没有来源的硬编码数字都是一个未记录的假设——审核者无法审核它。

## 通用工作流程

六个步骤，适用于每个非简单构建。

1. **打开/保存生命周期。** 开始时 Run `open`，结束时 Run `save` flush 到磁盘；`save` 只写入并保留 resident。仅在需要一次性交接时 Run `close`。大量 cell 操作使用 Batch：建议每块不超过 50 个 operation；纯 value-set payload 每块经测试可达 80+。cross-sheet formula 块是例外，应用独立 Batch（保留 `!`，见「已知问题」）。只在非 Office 工具边界 flush。
2. **创建或了解现状。** 新建 Run `create`；已有文件先 Run `view outline`。
3. **增量构建。** 每次结构操作（新增 sheet、chart、named range、pivot）后，先在目标上 Run `get` 确认结构，再追加更多内容。
4. **格式化。** 设置 column width、number format、freeze pane、tab color、header fill。formatting 是交付内容，不是可选润色。
5. **保存，并处理 cache。** `save` 写入磁盘。新建 formula 初始没有 cached value；人工在 spreadsheet app 中打开时会重算。下游 `INDEX/MATCH`、`SUMPRODUCT` 等会缓存写入时上游的 cached value（常为 `0`）；多 formula 构建后应重新触碰每个下游 cell（相同 formula 再 `set`）。之后对几个下游 cell Run `get`，确认 `cachedValue=` 合理。`validate` 会自行 flush pending edit。
6. **QA：假定存在问题。** 见 QA。最后一条命令 exit 0 并不代表完成。

## 快速入门

最小可行 xlsx：3 个月 revenue、一个 total formula、column width 和 currency format。请按实际文件和数据调整，不要直接复制粘贴。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": ["/workspace/book.xlsx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "open",
    "command_arguments": ["/workspace/book.xlsx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/book.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/Sheet1/A1",
        "props": { "value": "Month", "bold": true }
      },
      {
        "command": "set",
        "path": "/Sheet1/B1",
        "props": { "value": "Revenue", "bold": true }
      },
      {
        "command": "set",
        "path": "/Sheet1/A2",
        "props": { "value": "Jan" }
      },
      {
        "command": "set",
        "path": "/Sheet1/A3",
        "props": { "value": "Feb" }
      },
      {
        "command": "set",
        "path": "/Sheet1/A4",
        "props": { "value": "Mar" }
      },
      {
        "command": "set",
        "path": "/Sheet1/B2",
        "props": { "value": 42000, "numFmt": "$#,##0" }
      },
      {
        "command": "set",
        "path": "/Sheet1/B3",
        "props": { "value": 45000, "numFmt": "$#,##0" }
      },
      {
        "command": "set",
        "path": "/Sheet1/B4",
        "props": { "value": 48000, "numFmt": "$#,##0" }
      },
      {
        "command": "set",
        "path": "/Sheet1/A5",
        "props": { "value": "Total", "bold": true }
      },
      {
        "command": "set",
        "path": "/Sheet1/B5",
        "props": {
          "formula": "SUM(B2:B4)",
          "bold": true,
          "numFmt": "$#,##0"
        }
      },
      {
        "command": "set",
        "path": "/Sheet1/col[A]",
        "props": { "width": 12 }
      },
      {
        "command": "set",
        "path": "/Sheet1/col[B]",
        "props": { "width": 15 }
      }
    ],
    "stop_on_error": true
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "close",
    "command_arguments": ["/workspace/book.xlsx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": ["/workspace/book.xlsx"]
  }
}
```

已验证：`validate` 返回 `no errors found`，`B5` 会解析为 `135000`。基本流程是：打开 → set cell/formula → format → close → validate。

## CSV / 批量导入

**原生 `import`（CSV/TSV 首选）。** 一次调用将 CSV 加载到 sheet。`header` 会在 row 1 设置 auto-filter 与 freeze pane。column width 与 `numFmt` 仍需后续处理（遵循 dashboard Skill 的 D-12）。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "import",
    "command_arguments": [
      "/workspace/book.xlsx",
      "/Sheet1",
      "--file", "data.csv",
      "--header"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "import",
    "command_arguments": [
      "/workspace/book.xlsx",
      "/Sheet1",
      "--file", "data.tsv",
      "--format", "tsv",
      "--header"
    ]
  }
}
```

**分块 Batch fallback**（对应原 CLI 的 Python+batch 路径意图）：需要自定义 type coercion、formula injection，或 CSV 位于其他 data pipeline 时使用。适用于 600–6000+ cell：将 value-set 拆成每块约 80 个 `set` 的 Batch（任一块失败则降至 40）。number type inference 与 formula 可在后续通过目标 `set` 处理；本 recipe 的 batch 仅注入 value。结果量级参考：648 行 retail CSV（6490 个 cell）约 30 秒、零失败。

## 阅读与分析

先宽后窄。`outline` 先展示有哪些 sheet 与数据位置；明确目标后，再使用 `view` / `get` / `query` 深入查看。

**通过渲染后的 workbook 检查自己的结果。**

- Run `view html`：读取返回 HTML，审核渲染结果。每个 sheet 都可寻址，chart 会 inline render；可发现 `###`、placeholder 泄漏、pivot layout 与 row-height clip。
- `watch` 保留 live preview，供人工用户按需打开；agent 自检使用 `view html`。

每次 batch edit 后，应先用 `view html` 做首次视觉检查，并立即在源头修复。最终视觉验证应由用户在 Excel / WPS / Numbers 中打开 `.xlsx` 完成。

**了解概览。**

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": ["/workspace/book.xlsx", "outline"]
  }
}
```

**摘录。** 用于 content QA 的 plain-text dump；大文件可用 `start` / `end` / `cols` 限定范围。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": [
      "/workspace/book.xlsx",
      "text",
      "--start", "1",
      "--end", "50",
      "--cols", "A,B,C"
    ]
  }
}
```

其他 `view` 模式：`annotated`（cell 值 + 类型/formulas + 警告）、`stats`（数字摘要）、`issues`（损坏的 formulas、空 sheets、缺少引用）。

**Round-trip dump。** `dump` 将 workbook 或单个 sheet 序列化为可 replay 的 batch JSON；再用 Batch 重放。应使用它理解既有 workbook structure 或 clone/adapt template，而不是读取 raw OOXML。subtree dump 不携带 workbook-level resource（setting、named range），replay target 必须预先定义它们。完整范围见 Help / dump 说明。

**检查单个 element。** path 原样传入（含 `[N]`）。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": ["/workspace/book.xlsx", "/Sheet1/A1"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": ["/workspace/book.xlsx", "/Sheet1/A1:D10"]
  }
}
```

添加 `depth N` 可展开 child。完整 element list 见 Help `xlsx`。

**跨 workbook 查询。**

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": ["/workspace/book.xlsx", "cell:has(formula)"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": ["/workspace/book.xlsx", "cell:contains(\"#REF!\")"]
  }
}
```

运算符：`=`、`!=`、`~=`（包含）、`>=`、`<=`、`[attr]`（存在）。

**Merged cell shortcut。** `query` 的 `merge` 或 `mergedrange` 都是 `mergeCell` 的 alias。

**当数据量过大、逐行检查不再有效时，**使用 Excel 自己的 analysis element：

- `add type pivottable` 做分组/聚合；可附加 `slicer`。
- row 中 `sparkline`：`type` 为严格 enum **`line | column | stacked`**（`winloss` / `win-loss` 是 `stacked` 的 alias）。无效 `type=` 会 hard fail。
- 确切 prop name 见 Help `pivottable` / `slicer` / `sparkline`。

## 创建和编辑

一次构建的 90% 是 cell、formula、formatting 和一两个 chart。可用 verb：`add`、`set`、`remove`、`move`、`swap`；多项独立 mutation 用 Batch。

### Cell 与 formula

在一次调用中设置 value 及其 format。formula 开头不要写 `=`，CLI 会将其剥离。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/book.xlsx",
      "/Sheet1/B5",
      "--prop", "formula=SUM(B2:B4)",
      "--prop", "numFmt=$#,##0"
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
      "/workspace/book.xlsx",
      "/Sheet1/C5",
      "--prop", "formula=B5/A5",
      "--prop", "numFmt=0.0%"
    ]
  }
}
```

structure property（width、height、freeze、tabColor）位于 row / col / sheet node：

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/book.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/Sheet1/col[A]",
        "props": { "width": 20 }
      },
      {
        "command": "set",
        "path": "/Sheet1/row[1]",
        "props": { "height": 22 }
      },
      {
        "command": "set",
        "path": "/Sheet1",
        "props": { "freeze": "A2", "tabColor": "1F4E79" }
      }
    ],
    "stop_on_error": true
  }
}
```

### Named range

formula 中优先使用 named range 而非 `$B$6`。`ref` 同时含 `!` 与 `$`，用 Batch 写入（保留 JSON 字符串内的 `!`/`$`）：

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/book.xlsx",
      "/",
      "--type",
      "namedrange",
      "--prop",
      "name=GrowthRate",
      "--prop",
      "ref=Sheet1!$B$6"
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
      "/workspace/book.xlsx",
      "/Sheet1/B7",
      "--prop",
      "formula=B5*(1+GrowthRate)"
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
      "/workspace/book.xlsx",
      "/Sheet1/B6",
      "--prop",
      "value=0.05",
      "--prop",
      "numFmt=0.0%",
      "--prop",
      "font.color=0000FF"
    ]
  }
}
```

完整 schema 见 Help `namedrange`。

**batch JSON 不接受 shell alias。** 在 batch `props` 中始终使用完整 dotted name：`"font.color": "FF0000"`、`"font.size": 14`；不要使用 `"color": "FF0000"`。规则：任何 batch JSON 或 cell prop 都明确写为 `font.color` / `fill`。workbook-level 的 `parent` 必须是 `"/"`，sheet-scoped 的则为 `"/SheetName"`；空字符串不等效。

### 图表

chart type 见 Help `chart`，enum 很长（20+）。应根据所传达的信息选择：`column` 用于 category comparison、`line` 用于 time series、`pie` 仅用于 slice 比例一目了然的情况、`scatter` 用于 correlation。不要使用花哨 type，除非它确实回答特定问题。

**提供 chart data 有三种方式。每个 chart 只选择一种；在 `add` 时混用是常见陷阱。**

| 方式 | 形式 | 使用场景 |
|---|---|---|
| (a) inline `data` | `data="Sales:100,200,300"` + `categories="Jan,Feb,Mar"` | 小型演示 chart；数字不再可编辑 |
| (b) 2-D `dataRange` | `dataRange="Sheet1!A1:B4"` | 常规场景；必须是 **2-D** |
| (c) per-series point | `series1.name` / `series1.values` / `series1.categories` | multi-series 或不连续 range |

**Single-column 陷阱。** `dataRange="Sheet1!B2:B13"` 会被拒绝。应扩大 range 含 category column，或改用方式 (c)。

**创建后移动/调整 chart：**`set chart[N]` 的 `anchor` / `x` / `y` / `width` / `height`。**series 仍不可变**；要改 series 需 `remove` 后完整 `add`。`remove chart[1]` 会导致后续 index 下移；重添加附加在末尾。

**Anchor 尺寸。** 无 auto-fit。含 5–6 个 category 与 2 个 series 的 `column` chart，大约需要 `A5:L22`。拿不准时先从较小尺寸开始，通过 `view html` 后逐步放大。

**chart `dataRange` 必须带 sheet prefix。** 即使同 sheet 也写 `Summary!A17:C22`。

扩展 chart type（`boxWhisker`、`waterfall`、`funnel`、`histogram`、`treemap`、`sunburst`、`pareto`）仅在 data 确有需要时使用。

**不得在 chart title / series name / legend / axis title 中保留未替换的 template token。**

### 条件格式

三种常见类型（见 Help `cf`）：

- **Color scale**：`type=colorscale` 与 `minColor` / `midColor` / `maxColor`。
- **Data bar**：`type=databar`；显式 `min` / `max` 可使 column 内 scale 一致。
- **Formula rule**（`formulacf`）：`type=formula`、`formula="$C2>1000"` 和 fill/font。

规则：谨慎使用 CF。每个 cell 都有颜色的 workbook 不会传达有效信息。

### 数据验证

tracker 与 template 中的 input cell 必须使用 data validation。

**(a) Inline list：**

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/book.xlsx",
      "/Sheet1",
      "--type", "validation",
      "--prop", "sqref=C2:C100",
      "--prop", "type=list",
      "--prop", "formula1=Yes,No,Maybe",
      "--prop", "showError=true",
      "--prop", "errorTitle=Invalid",
      "--prop", "error=Select from list"
    ]
  }
}
```

**(b) Named range（cross-sheet lookup 首选）：**

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/book.xlsx",
      "/",
      "--type",
      "namedrange",
      "--prop",
      "name=StatusList",
      "--prop",
      "ref=Lookups!$A$2:$A$4"
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
      "/workspace/book.xlsx",
      "/Sheet1",
      "--type",
      "validation",
      "--prop",
      "sqref=B2:B100",
      "--prop",
      "type=list",
      "--prop",
      "formula1==StatusList"
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
      "/workspace/book.xlsx",
      "/Sheet1/B1",
      "--prop",
      "value=Status",
      "--prop",
      "bold=true"
    ]
  }
}
```

**(c) Direct cross-sheet range：**在 `formula1` 中直接写 `Lookups!$A$2:$A$4`（Batch JSON 保留 `!`/`$`）。写入后 Run `get` 验证：`formula1=` 必须显示未带反斜杠的 `!`。

其他常见 `type`：`decimal`、`whole`、`date`、`textLength`、`custom`。见 Help `validation`。

### 其他 element（简表）

- **Table**（ListObject）：`add type table`，见 Help `table`。
- **Comment**：`add type comment`，用于记录 hardcoded assumption。
- **Sheet reorder**：使用 `move`，不要使用 `swap`；`swap` 仅适用于 row/cell path。

## 按 role 操作 chart axis

按 **role**（`value` = Y，`category` = X）而非 index 访问 axis；XML order 不稳定。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": [
      "/workspace/book.xlsx",
      "/Sheet1/chart[1]/axis[@role=value]"
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
      "/workspace/book.xlsx",
      "/Sheet1/chart[1]/axis[@role=value]",
      "--prop", "min=0",
      "--prop", "max=100000"
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
      "/workspace/book.xlsx",
      "/Sheet1/chart[1]/axis[@role=category]",
      "--prop", "title=Month"
    ]
  }
}
```

安全 props：`title`、`min`、`max`、`majorGridlines`、`visible`、`labelRotation`。

## QA（必须执行）

**假定存在问题；你的任务是找到它们。**

第一次生成的 workbook 几乎不会完全正确。将 QA 当作 bug hunt，而不是确认步骤；首次检查发现零问题通常表示检查不够仔细。

### 声明“完成”前的最小循环

1. Run `view issues`：检查 empty sheet、broken formula、missing reference。
2. Run `view annotated`（sample range）：检查 value + type + warning。
3. 对每个 Excel error type 执行 query：`#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#N/A`。
4. Run `validate`：会自行 flush pending edit。
5. **视觉检查：** Run `view html` 并读取返回的 HTML path。检查 `###`、截断的 title、placeholder token、裁切 chart、纯白 pie slice 与空 chart anchor。`validate` 通过不等于可以交付。人工预览可 `watch` 或在 Excel / WPS / Numbers 中打开。
6. **Print layout 修复（宽 table / 多 chart sheet）。** 短 summary：`fitToPage=true`；高数据表：`fitToPage=1x0`（见上文 Print layout）。
7. 发现任一问题后先修复，再**重新运行完整循环**。

`view issues` + `view html` 构成结构 QA 组合。chart fill color / theme tint 可能因 viewer 而异；color fidelity 很重要时，应在用户的 target viewer 中 spot-check。

### Formula 验证清单

- [ ] 随机选取 2–3 个 formula，并在每个上 Run `get`。确认 formula string 符合预期，且 `cachedValue=` 与心算结果一致。
- [ ] **每个 summary cell 的 cached value 合理。** 若 progress tracker 在空 template 上显示 `199 / 199 / 100%`，说明 cache 错误；重新触碰 formula 或手动修正。
- [ ] **每个 number column 抽查一个 cell。**
- [ ] range 必须包含每一 row：数据到达 `B13` 时仍使用 `SUM(B2:B12)` 是最常见错误。
- [ ] cross-sheet formula 不得包含 `\!`。若 `get` 显示 `Sheet1\!A1`，删除后通过 Batch 重写。
- [ ] named range 必须指向其名称所表述的内容。
- [ ] 每一个 `/` denominator 都应受到保护：`IFERROR(x/y, 0)` 或 `IF(y=0, 0, x/y)`。
- [ ] chart data 要与 source cell 一致。
- [ ] chart title / series name / legend 不得包含未替换 token。

### Template QA

编辑 template 时，检查残留 placeholder：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": ["/workspace/book.xlsx", "cell:contains(\"{{\")"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": ["/workspace/book.xlsx", "cell:contains(\"xxxx\")"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": ["/workspace/book.xlsx", "cell:contains(\"TBD\")"]
  }
}
```

### 新鲜视角

完成 workbook 后，应重新打开它。像新 reviewer 一样从上到下阅读 `view text` / HTML preview，寻找 formula 错误、异常数字、format 不一致与缺失数据。

### 已知边界

`validate` 能捕获 schema error，不能捕获 design error。即使每个数字都错误，workbook 仍可能通过 `validate`。上述 checklist，尤其是用 source cell 抽查 formula，是发现 validation 未覆盖问题的方法。

## 已知问题和陷阱

### Cross-sheet `!` 陷阱（简述）

在 shell 路径下，`Sheet1!A1` 中的 `!` 可能被改写为 `\!`。Agent 路径用 JSON 字符串传入，避免 shell 层；但仍须验证。

**修复。** 用 Batch 写入含 `!` 的 formula：

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/book.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/Summary/B2",
        "props": { "formula": "Revenue!B13" }
      },
      {
        "command": "set",
        "path": "/Summary/B3",
        "props": { "formula": "Revenue!B14" }
      },
      {
        "command": "set",
        "path": "/Summary/B4",
        "props": { "formula": "B2+B3" }
      }
    ],
    "stop_on_error": true
  }
}
```

**验证。** 写入后对该 cell Run `get`；`formula=` 必须显示未带反斜杠的 `!`。

### CLI bug backlog（简述）

- **chart series 在创建后不可变：**改 series 需完整 `remove` + `add`。position 可用 `set`。
- **跨 sheet formula batch 在 resident 中运行良好。** 纯 value-set batch 在 50–80+ 操作下保持可靠。多 resident 竞争仍可能挂起。
- **conditional formatting 命名不对称：**element `conditionalformatting`，path `/cf[N]`。
- **`add` table 的 `position` prop：**常被忽略；用 `move index` / `after` / `before`。
- **`remove /sheet[N]` 的 cascade guard：**有 dependent 时先移除 dependent。
- **batch JSON 拒绝 cell `color` alias：**始终 `"font.color"` / `"fill"`。

### Renderer 注意事项（跨 viewer 的 color fidelity）

`view html` 是 structure QA 的正确工具。部分 chart render detail 取决于最终 viewer：

- 部分 viewer 中 pie / doughnut fill 塌缩成单一 theme tint。
- 部分 viewer 中 series color 偏离 theme。
- form-control checkbox 可能双框。

判定 color「损坏」前，在用户 target viewer 中打开。若那里正常，属于 viewer render，不应继续追逐。structure check（`###`、截断、placeholder、layout）仍是权威依据。

### Escape layer

`$` / `!` 在 Agent 路径写在 JSON 字符串内。另有两层：

- **JSON layer（batch）。** 标准 JSON escape：`"\n"`、`"\t"`、`"\""`。
- **Excel layer。** cell 内 `\n` 是实际 newline，配合 `wrapText=true`。拿不准时用 `get` 逐字符比较。

### 其他常见陷阱

| 陷阱 | 修复 |
|---|---|
| 猜测 prop name | 使用 Help `xlsx <element>` |
| cell 上的 `color=...` | 使用 `font.color` 或 `fill`；batch 用完整 dotted name |
| hex color `#FF0000` | 去掉 `#`：`FF0000` |
| `index` 与 `[N]` | `index` 从 0；`[N]` path 从 1 |
| 含空格的 sheet name | path 写 `"/My Sheet/A1"` |
| 年份显示为 `2,026` | `type=string` 或 `numFmt="@"` |
| 修改仍在 Excel 中打开的文件 | 先关闭 Excel |
| `swap` 不会重新排序 sheet | 使用 `move after` / `before` / `index` |
| 写入后 cached value 缺失 | 人工打开后才填充；`validate` 接受该状态 |

### Help 速查

疑问时用 `{{OFFICE_HELP_TOOL}}`（arguments：`xlsx` / element / verb+element）。Help 是权威 schema；本 Skill 帮助决策。
