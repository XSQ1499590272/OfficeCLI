---
name: officecli-xlsx
description: "任何涉及 .xlsx 文件的场景都使用此 Skill，包括创建电子表格、财务模型、仪表盘或追踪表；读取、解析或提取 .xlsx 数据；编辑、修改或更新既有工作簿；处理公式、图表、透视表或模板；将 CSV/TSV 数据导入 Excel 格式。当用户提到 'spreadsheet'、'workbook'、'Excel'、'financial model'、'tracker'、'dashboard' 或 .xlsx/.csv 文件名时触发。"
---

# OfficeCLI XLSX Skill

## ⚠️ Help 优先规则

**本 Skill 说明高质量 xlsx 应达到的标准，而不是罗列每一个 command flag。遇到不确定的 property name、enum value 或 alias 时，先查 Help，切勿猜测。**

```bash
officecli help xlsx                         # List all xlsx elements
officecli help xlsx <element>               # Full element schema (e.g. pivottable, chart, cf)
officecli help xlsx <verb> <element>        # Verb-scoped (e.g. add chart, set cell)
officecli help xlsx <element> --json        # Machine-readable schema
```

Help 与已安装的 CLI 版本一致。如本 Skill 与 Help 不一致，**以 Help 为准**。

## Shell 与执行规范

**Shell quoting（zsh / bash）。** Excel path 包含 `[]`，number format 可能包含 `$`；两者均为 shell 元字符。规则如下：

- 始终引用 element paths：`"/Sheet1/row[1]"`，而不是 `/Sheet1/row[1]`。
- 任何含 `$` 的 prop value（例如 `numFmt='$#,##0'`）都应使用**单引号**。
- formula 含跨 sheet `!` reference 时，使用 `batch` 加 `<<'EOF'` heredoc（见“已知问题”）。
- CLI 会解释 prop value 中的 `\n` 与 `\t`：`\n` 为真实 cell 内换行（与 `--prop wrapText=true` 配合），`\t` 为 tab。xlsx / docx / pptx 的行为一致。若需字面量反斜杠加 n，使用 `\\n`（通常不需要）。`$` 属于前述 shell 层问题，须使用单引号。

**增量执行。** 每次只运行一条 command 并检查 exit code。`officecli` 每次调用都会修改文件；含 50 条 command 的脚本若在第 3 条失败，后续操作会静默连锁失败。正确节奏是：一条 command → 检查 output → 继续。

## 输出标准

执行 command 前，先理解高质量 xlsx 的定义。以下是每个 workbook 必须达到的交付标准。

### 所有 Excel 文件

**formula error 必须为零。** 每个交付的 workbook 必须没有 `#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#N/A`。没有例外；使用 `IFERROR` 或 `IF(x=0,...)` 保护分母。

**应使用 formula，而非 hardcoded value。** 数字若可由其他 cell 计算得出，就应写为 formula。在应使用 `=SUM(B2:B9)` 的位置硬编码 `5000`，会破坏 workbook 随输入变化而保持 live 的约定。这是本 Skill 最重要的一条规则。

**专业 font。** workbook 中应统一使用一种专业 font（Arial / Calibri / Times New Roman）。不要因为某个 sheet 来自 CSV 就混用四种 font。

**显式 column width。** 不存在 auto-fit。用户会阅读的 column 必须设置 `width`；默认 8.43 个字符会截断内容。合理起点：label 20–25、number 12–15、date 12、short code 8–10。

**保留现有 template。** 编辑已有视觉样式的文件时应遵循原有约定；它们优先于本指南。

### 视觉交付底线（适用于每个 workbook）

声明完成前，运行 `officecli view "$FILE" html` 并读取返回的 HTML path，确认以下全部条件：

- **任何 cell 中不得出现 `###`。** `###` 表示 column 宽度不足以展示最长 value。用户会阅读的每个 column 都需要显式 `width`。交付文件中的 `###` 是未完成工作，不是“小的视觉瑕疵”。
- **不得截断 title。** sheet title、section header、long label 都必须完整显示；加宽 column 或在 cell 上设置 `wrapText=true`。
- **不得将 placeholder token 渲染为数据。** `$fy$24`、`{var}`、`<TODO>`、`xxxx` 不得出现在 cell、chart title、series name 或 legend 中；它们是未被替换的 build-time token。
- **pie / doughnut slice 应使用不同 fill color。** 若 slice 渲染成同色，切换为 `bar` / `column`，或显式设置 `colors=...`。
- **不得存在空尾页或空 chart anchor。** 例如，源 cell 为空时的 `anchor=D2:J18` 看起来像损坏的 chart。

如果上述任何一项失败，请在宣布完成之前停止并修复。

**Print layout。** 用户可能打印或作为 board pack 发送的任何 sheet 都需要 page setup。默认 portrait 且不 fit-to-page 会把宽 table 和 chart 拆到多页。应按 sheet 设置：

```bash
officecli set "$FILE" "/Summary" --prop orientation=landscape --prop fitToPage=true
```

触发条件：sheet 含 chart、超过 8 个 column，或用户需求提到 print / board / investor。

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

reviewer 应在读取 formula 前，仅通过 color 判断 cell 的类型。这是沟通约定，不是外观偏好。

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

1. **打开/保存生命周期。** 开始时使用 `officecli open <file>`，结束时使用 `officecli save <file>` flush 到磁盘；`save` 只写入并保留 resident 以供后续编辑。仅在需要一次性交接时使用 `officecli close <file>` 释放 resident。两者始终安全，不会报错或丢失工作。大量 cell 操作使用 `batch`：建议 **每块不超过 50 个 operation；纯 value-set payload 每块经测试可达 80+ 个 operation 且无失败。cross-sheet formula batch 是例外，应在 non-resident 的单一 heredoc 中运行（见“已知问题”）**。**只在非 officecli 边界 flush：**officecli 自身读取始终能看到编辑；仅在非 OfficeCLI 程序（openpyxl/pandas、Excel、renderer、交付流程）读取文件前运行 `save`/`close`。
2. **创建或了解现状。** 新建使用 `officecli create "$FILE"`；已有文件先运行 `officecli view "$FILE" outline` 了解结构。
3. **增量构建。** 每次运行一条 command，读取 output 后再继续。每次结构操作（新增 sheet、chart、named range、pivot）后，先在目标上运行 `get` 确认结构，再追加更多内容。
4. **格式化。** 设置 column width、number format、freeze pane、tab color、header fill。根据“输出标准”，formatting 是交付内容，不是可选润色。
5. **保存，并处理 cache。** `officecli save <file>` 会写入磁盘。新建 formula 初始没有 cached value；人工在 spreadsheet app 中打开文件时，应用会重新计算并填充它。**但下游 `INDEX/MATCH`、`SUMPRODUCT` 或任何引用上游 formula 的 formula，会缓存写入时上游 formula 的 cached value（常为 `0` 或过期值）；这个错误 cache 会保留在不重新计算的 reader 中。** 每次涉及 array formula（如带动态条件的 `SUMPRODUCT`、`SUMIFS`）或 cross-sheet chain 的多 formula 构建后，都应**重新触碰每个下游 cell**：用相同 formula 再运行一次 `set`，使 engine 基于新 cache 的上游重算其 cache。⚠️ 通过 resident 对 cross-sheet chain 重新触碰并不可靠（见 batch / resident 注意事项），更适合使用 non-resident `set` 进行该步骤。之后对几个下游 cell 执行 `officecli get`，确认 `cachedValue=` 合理。打开 resident 时使用 `validate` 是安全的，且它会自行将 pending edit flush 到磁盘（与 docx / pptx 相同）。
6. **QA：假定存在问题。** 见 QA。最后一条 command 的 exit code 为 0 并不代表完成；必须完成一次修复与验证循环，并且没有发现新问题。

## 快速入门

最小可行 xlsx：3 个月 revenue、一个 total formula、column width 和 currency format。请按实际文件和数据调整，不要直接复制粘贴。

```bash
officecli create "$FILE"
officecli open "$FILE"
officecli set "$FILE" /Sheet1/A1 --prop value=Month --prop bold=true
officecli set "$FILE" /Sheet1/B1 --prop value=Revenue --prop bold=true
officecli set "$FILE" /Sheet1/A2 --prop value=Jan
officecli set "$FILE" /Sheet1/A3 --prop value=Feb
officecli set "$FILE" /Sheet1/A4 --prop value=Mar
officecli set "$FILE" /Sheet1/B2 --prop value=42000 --prop numFmt='$#,##0'
officecli set "$FILE" /Sheet1/B3 --prop value=45000 --prop numFmt='$#,##0'
officecli set "$FILE" /Sheet1/B4 --prop value=48000 --prop numFmt='$#,##0'
officecli set "$FILE" /Sheet1/A5 --prop value=Total --prop bold=true
officecli set "$FILE" /Sheet1/B5 --prop formula="SUM(B2:B4)" --prop bold=true --prop numFmt='$#,##0'
officecli set "$FILE" "/Sheet1/col[A]" --prop width=12
officecli set "$FILE" "/Sheet1/col[B]" --prop width=15
officecli close "$FILE"
officecli validate "$FILE"
```

已验证：`validate` 返回 `no errors found`，`B5` 会解析为 `135000`。基本流程是：打开 → set cell/formula → format → close → validate。

## CSV / 批量导入

**原生 `import` command（CSV/TSV 首选）。** 这是最快的 path，一次调用即可将 CSV 加载到 sheet。`--header` 会在 row 1 设置 auto-filter 与 freeze pane。column width 与 `numFmt` 仍需后续处理（遵循 dashboard Skill 的 D-12）。

```bash
officecli import "$FILE" /Sheet1 --file data.csv --header
officecli import "$FILE" /Sheet1 --file data.tsv --format tsv --header
officecli import "$FILE" /Sheet1 --stdin --start-cell B2 < data.csv
```

**Python + batch fallback**：需要自定义 type coercion、formula injection，或 CSV 位于其他 data pipeline 中时使用。适用于 600–6000+ cell：

```python
# gen_batch.py — produces batch chunks of 80 value-set ops each
import csv, json
ops = []
with open("data.csv") as f:
    reader = csv.reader(f)
    for r, row in enumerate(reader, start=1):
        for c, val in enumerate(row):
            col = chr(ord('A') + c)
            ops.append({"command":"set","path":f"/Data/{col}{r}",
                        "props":{"value": val}})
for i in range(0, len(ops), 80):
    print(json.dumps(ops[i:i+80]))
```

```bash
python gen_batch.py | while IFS= read -r chunk; do
  printf '%s\n' "$chunk" | officecli batch "$FILE"
done
```

结果：648 行 retail CSV（6490 个 cell）约 30 秒加载完成，零失败。建议从每块 80 个 operation 开始，任一块失败则降至 40。number type inference 与 formula 可在后续通过目标 `set` 处理；本 recipe 的 batch 仅注入 value。

## 阅读与分析

先宽后窄。`outline` 先展示有哪些 sheet 与数据位置；明确目标后，再使用 `view` / `get` / `query` 深入查看。

**通过渲染后的 workbook 检查自己的结果。**

- `officecli view $FILE html`：读取返回 HTML，审核渲染结果。每个 sheet 都可寻址，chart 会 inline render；可发现 `###`、placeholder 泄漏、pivot layout 与 row-height clip。
- `officecli watch $FILE`：保留 live preview，供人工用户按需打开。用户需要 watch 时使用；agent 自检使用上方的 `view html`。

每次 batch edit 后，应先用 `view html` 做首次视觉检查，并立即在源头修复。最终视觉验证应由用户在 Excel / WPS / Numbers 中打开 `.xlsx` 完成。

**了解概览。** 查看 sheet、尺寸与 formula 数量。

```bash
officecli view "$FILE" outline
```

**摘录。** 用于 content QA 或 LLM context 的 plain-text dump；大文件可用 `--start` / `--end` / `--cols` 限定范围。

```bash
officecli view "$FILE" text --start 1 --end 50 --cols A,B,C
```

其他值得了解的 `view` 模式：`annotated`（cell 值 + 类型/formulas + 警告）、`stats`（数字摘要）、`issues`（损坏的 formulas、空 sheets、缺少引用）。

**Round-trip dump。** `officecli dump "$FILE" [path]` 会将 workbook 或单个 sheet（`/Sheet1`、`/sheet[N]`）序列化为可 replay 的 batch JSON；`officecli batch new.xlsx --input dump.json` 可重放它。应使用它理解既有 workbook structure 或 clone/adapt template，而不是读取 raw OOXML。完整范围见 `dump --help`；subtree dump 不携带 workbook-level resource（setting、named range），replay target 必须预先定义它们。

```bash
officecli dump "$FILE" -o blueprint.json            # whole workbook
officecli dump "$FILE" /Sheet1 -o sheet.json        # one worksheet
officecli batch new.xlsx --input blueprint.json
```

**检查单个 element。** 使用 XPath-style path，并始终加引号，避免 shell glob `[N]`。

```bash
officecli get "$FILE" "/Sheet1/A1"            # one cell
officecli get "$FILE" "/Sheet1/A1:D10"        # range
officecli get "$FILE" "/Sheet1/chart[1]"      # chart
officecli get "$FILE" "/Sheet1/table[1]"      # ListObject
officecli get "$FILE" "/namedrange[1]"        # workbook-level named range
```

添加 `--depth N` 可展开 child；添加 `--json` 可获得 machine-readable output。完整 element list 见 `officecli help xlsx`。

**跨 workbook 查询。** 使用 CSS-like selector 做系统检查（formula coverage、error cell、empty title），而非手动遍历。

```bash
officecli query "$FILE" 'cell:has(formula)'       # every formula cell
officecli query "$FILE" 'cell:contains("#REF!")'  # broken references
officecli query "$FILE" 'cell[type=Number]'       # typed filter
officecli query "$FILE" 'Sheet1!B[value!=0]'      # sheet-scoped
```

运算符：`=`、`!=`、`~=`（包含）、`>=`、`<=`、`[attr]`（存在）。

**Merged cell shortcut。** `officecli query $FILE merge` 或 `mergedrange` 都是 `mergeCell` 的 alias，可返回 workbook 中每个 merged range，无需手动遍历 `<mergeCell>` entry。

**当数据量过大、逐行检查不再有效时，**使用 Excel 自己的 analysis element：

- 使用 `officecli add`（`--type pivottable`）构建 **pivot table** 做分组/聚合，无需编写 20 个 SUMIF；可附加 **slicer**（`--type slicer`）为读者提供 filter UI。
- 在 row 中使用 **sparkline**（`--type sparkline`）展示每行 trend，比每 row 一个 chart 更轻量，并可 inline print。`type` 为严格 enum：**`line | column | stacked`**（`winloss` / `win-loss` 是 `stacked` 的 alias）。无效 `type=` 会 hard fail，不再静默 fallback 到 `line`。
- 确切 prop name 见 `officecli help xlsx pivottable`、`officecli help xlsx slicer`、`officecli help xlsx sparkline`。

## 创建和编辑

一次构建的 90% 是 cell、formula、formatting 和一两个 chart。可用 verb：`add`（新增 element）、`set`（修改 prop）、`remove`、`move`、`swap`、`batch`。

### Cell 与 formula

在一次调用中设置 value 及其 format。formula 开头不要写 `=`，CLI 会将其剥离。

```bash
officecli set "$FILE" /Sheet1/B5 --prop formula="SUM(B2:B4)" --prop numFmt='$#,##0'
officecli set "$FILE" /Sheet1/C5 --prop formula="B5/A5" --prop numFmt="0.0%"
```

structure property（width、height、freeze、tabColor）位于 row / col / sheet node：

```bash
officecli set "$FILE" "/Sheet1/col[A]" --prop width=20
officecli set "$FILE" "/Sheet1/row[1]" --prop height=22
officecli set "$FILE" "/Sheet1" --prop freeze=A2 --prop tabColor=1F4E79
```

### Named range

formula 中优先使用 named range 而非 `$B$6`。它具备自说明性（`GrowthRate` 胜过 `$B$6`），也可以移动 assumption cell 而不破坏 formula。因为 `ref` value 同时包含 `!` 与 `$`，应通过 batch heredoc `add`：

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"GrowthRate","ref":"Sheet1!$B$6"}}
]
EOF
```

完整 schema 见 `officecli help xlsx namedrange`。

**batch JSON 不接受 shell alias。** 在 batch `props` 中始终使用完整 dotted name：`"font.color": "FF0000"`、`"font.size": 14`；不要使用 `"color": "FF0000"`，它无法区分 text 与 fill。普通 cell 的 shell 形式也会被拒绝：`--prop color=1F4E79` 会报 `ambiguous in cell context — use 'font.color' (text) or 'fill' (bg)`。规则是：任何 batch JSON 或 cell prop 都明确写为 `font.color` / `fill`。workbook-level 的 `parent` 必须是 `"/"`，sheet-scoped 的则为 `"/SheetName"`；空字符串不等效。

### 图表

chart type 见 `officecli help xlsx chart`，enum 很长（20+）。应根据所传达的信息选择：`column` 用于 category comparison、`line` 用于 time series、`pie` 仅用于 slice 比例一目了然的情况、`scatter` 用于 correlation。不要使用花哨 type，除非它确实回答特定问题。

**提供 chart data 有三种方式。每个 chart 只选择一种；在 `add` 时混用是常见陷阱。**

| 方式 | 形式 | 使用场景 |
|---|---|---|
| (a) inline `data` | `--prop data="Sales:100,200,300" --prop categories="Jan,Feb,Mar"` | 小型演示 chart；数字不再可编辑，source of truth 位于 chart XML 而不是 cell。 |
| (b) 2-D `dataRange` | `--prop dataRange="Sheet1!A1:B4"`（第一 column 为 category，第一 row 为 title / series name） | 常规场景。必须是 **2-D**；单一 column 会报 “Chart requires data”。 |
| (c) per-series point | `--prop series1.name=Sales --prop series1.values="Sheet1!B2:B4" --prop series1.categories="Sheet1!A2:A4"` | multi-series chart、每个 series 指向不连续 range，或需显式 series name。单独设置 `series1.values` 而不设置 `categories` 时，chart 的 x-axis 会显示 `1,2,3`。 |

**Single-column 陷阱。** `dataRange="Sheet1!B2:B13"` 看起来是“value column”，但 engine 会以 `Chart requires data` 拒绝它。应扩大 range 以包含 category column（`A2:B13`），或改用显式设置 `series1.categories` 的方式 (c)。

**创建后移动/调整 chart：**使用 `set chart[N] --prop anchor="F5:N25"`，也支持 `--prop x= --prop y= --prop width= --prop height=`。**series 仍不可变**；若要新增或修改 series，需要 `officecli remove` chart，再用完整 series list `officecli add`。注意 `remove chart[1]` 会导致 `chart[2] → chart[1]`，而重新添加会**附加在末尾**；要保留 chart order，应全部 remove 后按顺序重建。

**Anchor 尺寸。** 无 auto-fit。含 5–6 个 category 与 2 个 series 的 `column` chart，大约需要 `A5:L22`（12 column × 18 row）才能完整显示 label。过窄会截断 x-axis label，过宽则可能在 print/export 时跨页拆分。拿不准时先从较小尺寸开始，通过 `view html` 读取 HTML preview 后逐步放大；下方的 layout 设置是另一半解决方案。

**chart `dataRange` 必须带 sheet prefix。** 即使 chart 与 data 位于同一 sheet，也写 `dataRange="Summary!A17:C22"`，不要写 `A17:C22`。无 prefix 的写法行为不稳定；带 prefix 的形式可靠。

officecli 提供传统 Excel object model 缺少的扩展 chart type：`boxWhisker`、`waterfall`、`funnel`、`histogram`、`treemap`、`sunburst`、`pareto`。仅在 data 确有需要时使用。

**不得在 chart title / series name / legend / axis title 中保留未替换的 template token。** `$fy$24`、`{var}`、`<TODO>`、`$VAR`、`{{placeholder}}` 会在 legend 中**按字面量**渲染；即使 `validate` 通过，CFO 仍会看到本应显示 “FY2024” 的 `$fy$24`。必须绑定到最终 text 或 cell reference，例如 `title="FY2024 Revenue"` 或 `series1.name="Sheet1!A1"`。

### 条件格式

三种常见类型，各自有不同的 prop 结构（见 `officecli help xlsx cf`）：

- **Color scale**：cell 随 value 渐变着色，使用 `type=colorscale` 与 `minColor` / `midColor` / `maxColor`。
- **Data bar**：cell 内的 bar 展示 value 大小，使用 `type=databar`。显式设置 `min` / `max` 可使 column 内 scale 一致；省略时 default value 也有效。
- **Formula rule**（`formulacf` element）：condition 为 true 时突出显示 row，使用 `type=formula`、`formula="$C2>1000"` 和 fill/font。

规则：谨慎使用 CF。每个 cell 都有颜色的 workbook 不会传达有效信息。

### 数据验证

tracker 与 template 中的 input cell 必须使用 data validation。成本很低，却能阻断大量 downstream error。**列表 source 有三种模式**，按 allowed value 的位置选择。

**(a) Inline list：**allowed value 很少且固定在 rule 内。

```bash
officecli add "$FILE" /Sheet1 --type validation \
  --prop sqref="C2:C100" --prop type=list \
  --prop formula1="Yes,No,Maybe" \
  --prop showError=true --prop errorTitle="Invalid" --prop error="Select from list"
```

**(b) Named range（cross-sheet lookup 首选）：**allowed value 位于其他 sheet 且可能增长。先定义 named range，再引用它。`ref` 包含 `!` 和 `$`，请使用 batch heredoc：

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"StatusList","ref":"Lookups!$A$2:$A$4"}},
  {"command":"add","parent":"/Sheet1","type":"validation","props":{"sqref":"B2:B100","type":"list","formula1":"=StatusList"}}
]
EOF
```

**(c) Direct cross-sheet range：**在 `formula1` 中直接写 raw `Lookups!$A$2:$A$4`，不使用 named range。同样需要 batch heredoc 以保留 `!` 与 `$`：

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/Sheet1","type":"validation","props":{"sqref":"C2:C100","type":"list","formula1":"Lookups!$A$2:$A$4"}}
]
EOF
```

若在 shell 中将 cross-sheet variant 写成 `--prop formula1=...`，`!` 会被 shell 改写为 `\!`，下拉 list 会静默退化为空。使用 `officecli get "$FILE" /Sheet1/validation[N]` 验证：`formula1=` 必须显示未带反斜杠的 `!`。

其他常见 `type`：`decimal`、`whole`、`date`、`textLength`、`custom`。operator 与完整 prop list 见 `officecli help xlsx validation`。

### 其他 element（简表）

- **Table**（ListObject）：使用带 range 的 `add --type table`，可获得 auto-filter 与 structured reference。见 `officecli help xlsx table`。
- **Comment**：`add --type comment`，用于记录 hardcoded assumption。见 `officecli help xlsx comment`。
- **Sheet reorder**：使用 `officecli move`，不要使用 `swap`；`swap` 仅适用于 row/cell path。

## 按 role 操作 chart axis

直接编辑 chart axis 比重建 chart 更轻量。按**role**（`value` = Y，`category` = X）而非 index 访问 axis；XML order 不稳定。

```bash
officecli get "$FILE" "/Sheet1/chart[1]/axis[@role=value]"
officecli set "$FILE" "/Sheet1/chart[1]/axis[@role=value]" --prop min=0 --prop max=100000
officecli set "$FILE" "/Sheet1/chart[1]/axis[@role=category]" --prop title="Month"
```

安全props：`title`、`min`、`max`、`majorGridlines`、`visible`、`labelRotation`。

## QA（必须执行）

**假定存在问题；你的任务是找到它们。**

第一次生成的 workbook 几乎不会完全正确。将 QA 当作 bug hunt，而不是确认步骤；首次检查发现零问题通常表示检查不够仔细。formula 看起来正常，**直到**你用 source cell 抽查其中两个。

### 声明“完成”前的最小循环

1. `officecli view "$FILE" issues`：检查 empty sheet、broken formula、missing reference。
2. `officecli view "$FILE" annotated`（sample range）：检查 value + type + warning。
3. 对每个 Excel error type 执行 query：
   ```bash
   officecli query "$FILE" 'cell:contains("#REF!")'
   officecli query "$FILE" 'cell:contains("#DIV/0!")'
   officecli query "$FILE" 'cell:contains("#VALUE!")'
   officecli query "$FILE" 'cell:contains("#NAME?")'
   officecli query "$FILE" 'cell:contains("#N/A")'
   ```
4. `officecli validate "$FILE"`：打开 resident 后的安全检查；`validate` 会自行将待写入的编辑 flush 到磁盘。
5. **视觉检查：通过 HTML preview 检查每个 sheet。** 运行 `officecli view "$FILE" html` 并读取返回的 HTML path。每个 sheet 会将 chart inline render。检查 `###`、截断的 title、placeholder token（`$fy$24`、`{var}`、`<TODO>`）、裁切 chart、纯白 pie slice 与空 chart anchor；发现任一问题都应在声明完成前停止并修复。`validate` 通过不等于可以交付；交付目标是 preview 看起来像真实 workbook。人工预览可运行 `officecli watch "$FILE"`（用户按需打开 live preview），或直接在 Excel / WPS / Numbers 中打开 `.xlsx`。
6. **修复 print layout（宽 table / 多 chart sheet）。** sheet 含 chart 或宽 table 且用户将打印时，应设置每页 layout 使其 fit on one page：
   ```bash
   officecli set "$FILE" "/Summary" --prop orientation=landscape --prop fitToPage=true
   ```
结果是：每个 sheet 的 print layout 保持在一页内，不会在 chart 中间分割。适用于含 chart 或超过 8 column table 的每个 sheet。
7. 发现任一问题后先修复，再**重新运行完整循环**；一次修复常会引入另一个问题。

`officecli view issues` + `view html` 构成结构 QA 组合：`issues` 捕获 broken formula 与 empty sheet，`view html`（读取返回 HTML path）捕获 `###`、截断与 token 泄漏。chart fill color / theme tint 可能因 viewer 而异；color fidelity 很重要时，应在用户的 target viewer 中 spot-check。

### Formula 验证清单

- [ ] 随机选取 2–3 个 formula，并在每个上运行 `officecli get`。确认 formula string 符合预期，且 `cachedValue=` 与心算结果一致。
- [ ] **每个 summary cell 的 cached value 合理。** 任意 aggregate cell（COUNTA / COUNTIF / SUMPRODUCT / INDEX&MATCH）都必须有合理 `cachedValue`。若 progress tracker 在空 template 上显示 `199 / 199 / 100%`，说明 cache 错误；通过 `set` 重新触碰 formula 以强制重算，或手动 `set` 为正确 cached value。不能交付“`validate` 通过但数字虚构”的文件。
- [ ] **每个 number column 抽查一个 cell。** `%` column 总显示整数 `0.0%` 说明 denominator 错误或 numerator cache 过期；应调查该 cell 并修复模式。
- [ ] range 必须包含每一 row：数据到达 `B13` 时仍使用 `SUM(B2:B12)` 是最常见错误。
- [ ] cross-sheet formula（`Sheet1!A1`）不得包含 `\!`。若 `officecli get` 显示 `Sheet1\!A1`，说明 `!` 被 shell 损坏；删除后通过 batch/heredoc 重写。
- [ ] named range（`officecli get "$FILE" "/namedrange[1]"`）必须指向其名称所表述的内容。
- [ ] 每一个 `/` denominator 都应受到保护：`IFERROR(x/y, 0)` 或 `IF(y=0, 0, x/y)`。
- [ ] chart data 要与 source cell 一致：每个含 inline data 的 chart 都应对照 source cell 用 `officecli get` 抽查 data point。
- [ ] chart title / series name / legend 不得包含未替换 token（`$...$`、`{var}`、`<TODO>`）；用 `officecli get /Sheet1/chart[N]` 检查 chart。

### Template QA

编辑 template 时，检查残留 placeholder；它们看起来像正常内容，可能绕过 `validate`：

```bash
officecli query "$FILE" 'cell:contains("{{")'
officecli query "$FILE" 'cell:contains("xxxx")'
officecli query "$FILE" 'cell:contains("TBD")'
```

### 新鲜视角

完成 workbook 后，应重新打开它。像新 reviewer 一样从上到下阅读 `view text` / HTML preview，寻找 formula 错误、异常数字、format 不一致与缺失数据。

### 已知边界

`validate` 能捕获 schema error，不能捕获 design error。即使每个数字都错误，workbook 仍可能通过 `validate`。上述 checklist，尤其是用 source cell 抽查 formula，是发现 validation 未覆盖问题的方法。

## 已知问题和陷阱

### Cross-sheet `!` 陷阱（简述）

Shell（bash history expansion、zsh split）与 CLI arg 可能将 `Sheet1!A1` 中的 `!` 改写为 `\!`。含 `\!` 的 formula 已被悄然破坏：它会显示为 literal text，且不再 reference 任何内容。

**修复。** 使用单引号 delimiter（`<<'EOF'`）的 batch heredoc，以禁用所有 shell expansion：

```bash
cat <<'EOF' | officecli batch "$FILE"
[{"command":"set","path":"/Summary/B2","props":{"formula":"Revenue!B13"}}]
EOF
```

**验证。**写入后，对该 cell 执行 `officecli get`；`formula=` 必须显示未带反斜杠的 `!`。

### CLI bug backlog（简述）

以下是 CLI 需要解决的限制与缺口，并非输出文件本身的缺陷。

- **chart series 在创建后不可变：**新增或修改 series 需要完整执行 `remove` + `add`。（position 可以修改：`set chart[N] --prop anchor=` / `x/y/width/height`。）`remove chart[N]` 会下移后续 index，重新添加则附加在末尾。
- **跨 sheet formula batch 在 resident 中运行良好。** 之前“3–5 个操作也会死锁”的警告已不再复现。纯 value-set batch 在 50–80+ 个操作下也保持可靠。遇到问题时，请回退到非 resident 的大 batch 或单独 `set`。**同一文件/计算机上的多个 resident process 仍可能竞争：**若其他 agent/session 已持有该文件的 resident process，可能出现非确定性挂起。
- **conditional formatting 的命名不对称：**`--type` 的 element name 是 `conditionalformatting`，path suffix 是 `/cf[N]`。schema 参见 `officecli help xlsx conditionalformatting`，path 使用 `/cf[N]`。
- **`add` table 的 `position` prop：**Help 显示支持 `position`，但该 prop 常被忽略。创建 sheet 后使用 `officecli move --index` / `--after` / `--before` 排序。
- **`remove /sheet[N]` 的 cascade guard：**当其他 sheet 的 validation / conditional formatting / sparkline / hyperlink / named range 引用该 sheet 时，会拒绝 remove/rename。先移除这些 dependent element，再 remove sheet。
- **batch JSON 拒绝 cell `color` alias：**在 batch `props` 中，`"color": "FF0000"` 会报 `ambiguous in cell context — use 'font.color' (text) or 'fill' (bg)`。shell-level CLI 允许非 cell element 使用 `--prop color=...` / `--prop size=14` alias，但 cell 的 batch JSON 始终必须使用完整 dotted name：`"font.color"`、`"font.size"`、`"font.name"`。

### Renderer 注意事项（跨 viewer 的 color fidelity）

`officecli view html` 是处理 structure QA（overflow、truncate、placeholder leak、layout）的正确工具，应读取返回的 HTML path。部分 chart render detail 取决于最终用户使用的 viewer。已观察到的差异：

- **部分 viewer 中 pie / doughnut fill color 会塌缩成单一 theme tint**（slice 看起来“全白”或“全为一种颜色”）。在用户的 target viewer 中可能实际正常。
- **部分 viewer 中，chart / column chart series color 可能偏离 workbook theme**。
- **form-control checkbox 可能在部分 viewer 中 render 为双框**。

判定 color 或 chart “损坏”前，请在用户实际的 target viewer 中打开文件。若那里正常，问题属于 viewer render 而非 data，不应继续追逐。CLI 的 structure check（`###`、截断、placeholder text、layout）仍是权威依据。

### Escape layer（shell quoting 见上；以下为额外层）

`$` 属于 shell layer（使用单引号，见上）。prop value 中的 `\n` / `\t` 会被 CLI 解释为真实 newline / tab。另有两层：

- **JSON layer（batch）。** 使用标准 JSON escape：`"\n"`、`"\t"`、`"\""`。最终 string 中的真实反斜杠写为 `"\\\\"`。
- **Excel layer。** cell 内的 `\n` 是实际 newline，应配合 `--prop wrapText=true`，使 Excel 显示换行。在 shell-quoted prop 中可直接写入（`--prop value='a\nb'`）；batch JSON 中的 `"\n"` 效果相同。拿不准时，用 `officecli get` 与 cell 逐字符比较。

### 其他常见陷阱

| 陷阱 | 修复 |
|---|---|
| `--name "foo"` | 所有 property 都通过 `--prop`：`--prop name="foo"` |
| 猜测 prop name | 使用 `officecli help xlsx <element>`，不要即兴猜测 |
| cell 上的 `--prop color=...` | 有歧义；使用 `font.color`（text）或 `fill`（background）。batch JSON 也一样：始终使用完整 dotted name，不使用 shell alias |
| hex color `#FF0000` | 去掉 `#`：`FF0000` |
| `--index` 与 `[N]` | `--index` 从 0 开始（array）；`[N]` path 从 1 开始（XPath） |
| zsh/bash 中未加引号的 `[N]` | 每个 path 都加引号：`"/Sheet1/row[1]"` |
| 含空格的 sheet name | 为完整 path 加引号：`"/My Sheet/A1"` |
| 年份显示为 `2,026` | 使用 `--prop type=string` 或 `numFmt="@"` |
| 修改仍在 Excel 中打开的文件 | 先关闭 Excel |
| `swap` 不会重新排序 sheet | `swap` 适用于 row/cell；使用 `move --after` / `--before` / `--index` 排序 sheet |
| 写入后 cached value 缺失 | 人工在 spreadsheet app 中打开后，新的 formula 才会获得 cached value；`validate` 接受这种状态 |
