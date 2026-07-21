# Financial Model Agent Skill

Author: xiesq, 2026-07-15

**本 Skill 是 `xlsx`/`excel` 之上的 scene layer。** 继承 Help-First、视觉底线、CFO 四色（蓝输入 / 黑公式 / 绿跨表 / 黄假设底）、数字格式标准、assumption 纪律、chart data-feed、Delivery cycle、cache-drift、Known Issues（跨表 `!`、batch+resident、renderer）。本文件只加：三区架构、3 模型配方（3-statement / DCF / LBO）、sensitivity / scenario、财务函数模式、circular-reference 纪律、Gates 4–6。

写入后 Run readback，检查 `#REF!` / `#DIV/0!` / `#VALUE!` / `#N/A` / `#NAME?`。不要用硬编码结果替换公式。

## ⚠️ Help 优先规则

**本 Skill 说明 financial model 需要什么，而不是罗列全部 CLI flag。** 不确定的 prop / alias / enum，先查 Help。

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
      "cell"
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
      "namedrange"
    ]
  }
}
```

Skill 与 Help 冲突时，**以 Help 为准**。

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

### Reverse handoff

budget tracker / CSV dump / ops KPI / 无 forecast 的 cap table → xlsx 基座。仅当提到 3-statement / DCF / WACC / NPV / TV / LBO / debt schedule / MOIC / IRR / unit economics / ARR / sensitivity / scenario / pro forma 时用本 Skill。

## 宿主工具与执行规范

（对应原 CLI「Shell & Execution Discipline」；shell 引号改为 Tool JSON 传参。）

**Path 与特殊字符。** number format 含 `$`、跨表 formula 含 `!`：直接写在 JSON 字符串内。含 `&` / 空格的 sheet 名写成 `'P&L'!B3`。

**增量执行。** 按依赖链构建；链后 cache-refresh。生命周期 `create` / `open` / `save` / `close`。

**阶段 0 Batch/Run 契约（固定补丁，非业务加戏）。**

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。

## Core Principles（identity）

八个 delta：

1. **三区强制：** Inputs → Calc → Outputs。塌缩 = 不可审计。
2. **假设在单元格，不在公式常量。** `=B5*(1+Assumptions!GrowthRate)`，永不 `=B5*1.05`。
3. **每期 statements 平衡。** Assets−Liab−Equity=0；CF.EndingCash=BS.Cash。Gate 4 见 `IMBALANCED` 拒绝。
4. **Hardcodes 审计。** Calc sheets 零硬编码数字；Gate 6 计数。
5. **Sensitivity / scenario 一等公民。** 2 轴网格、INDEX/MATCH 下拉、或 Base/Upside/Downside 列。Excel Data Table 不可靠——只用手工网格。
6. **估值单元格 cached value load-bearing。** 空 / `#OCLI_NOTEVAL!` 对不重算读者是错数。Gate 5 抽查。
7. **Circularity 是设计选择。** 合法环（interest↔cash、revolver plug）用 `calc.iterate=true`。意外环是代数错误——不要用 iterate 糊弄。
8. **≥3 次使用的假设用 named range。** 声明未用 = 死装饰；Gate 6 抓。

## Three-zone architecture（硬规则）

| Zone | Sheet names | Tab color | Content | Hardcodes | Formulas |
|---|---|---|---|---|---|
| **Inputs** | Assumptions / Inputs / Drivers | Yellow `FFC000` | 增长、margin、tax、WACC… | Blue `0000FF` | 仅派生假设允许 |
| **Calc** | P&L / Balance Sheet / Cash Flow / DCF / Debt / ARR | Blue `4472C4` | 全部推导 | **Zero** | 同表黑 / 跨表绿 |
| **Outputs** | Summary / Dashboard / Sensitivity / Returns | Green `70AD47` | KPIs、网格、charts | 仅标签；数字硬编码→0 | 黑/绿 |

**Build order：** Assumptions → Calc 依赖链（3-stmt：IS→BS→CF；DCF：FCF→WACC→NPV；LBO：S&U→Debt→…）→ Outputs 最后。先建 Outputs 会缓存一堆 0。

Executable zone audit（Run query）：Calc 上 `formula==null` 的 Number 单元格必须为 0；Assumptions 硬编码 drivers 应 ≥5。

## Print delivery（board / IC / LP）

触发 print / 一页 / 董事会 / 投资人 / IC memo / LP update：Print_Area 限定 Outputs；隐藏非 Outputs；fit-to-page landscape。需要 raw XML 时：prose + **approval** + 逐步 Run（**relationship**/path），**永不**进 Batch JSON。删掉 Calc 上冲突的 Print_Area。

## Build-order & cache-drift（3-statement 关键）

三事实造成静默错数：(1) 新公式无 cache；(2) 同序列写下游会吃到上游 pre-cache 0；(3) resident 下跨表 batch 可能卡死。

纪律：按数据链构建；链完成后 **cache-refresh pass**（重设 summary/估值/balance-check）；`get ... cachedValue` 应为合理非 null。见 `#OCLI_NOTEVAL!` → close residents，重设。

## Recipes — three model types

路径 `/workspace/model.xlsx`。先 create/open；结束 close + validate。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": ["/workspace/model.xlsx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "open",
    "command_arguments": ["/workspace/model.xlsx"]
  }
}
```

### Recipe A — 3-statement（P&L + BS + CF）

**产出：** Assumptions、P&L、Balance Sheet、Cash Flow、Summary。年份列 2024A·2025E·2026E·2027E。BS balance-check；CF cash-recon。

**强制顺序：** Assumptions → P&L → BS → CF → Summary。勿先建 BS（RE 依赖 NI）；勿先建 CF（Y1 cash 锚定 BS）。

**Step 1 — sheets + tab colors**

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/model.xlsx",
    "operations": [
      {
        "command": "add",
        "parent": "/",
        "type": "sheet",
        "props": { "name": "Assumptions", "tabColor": "FFC000" }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "sheet",
        "props": { "name": "P&L", "tabColor": "4472C4" }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "sheet",
        "props": { "name": "Balance Sheet", "tabColor": "4472C4" }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "sheet",
        "props": { "name": "Cash Flow", "tabColor": "4472C4" }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "sheet",
        "props": { "name": "Summary", "tabColor": "70AD47" }
      }
    ],
    "stop_on_error": true
  }
}
```

再 Batch/Run 设 freeze（Assumptions `B2`；P&L/BS/CF `B3`）。

**Step 2 — assumptions（蓝字 + 黄底关键驱动）**

Drivers：RevenueGrowth、GrossMargin、OpExRatio、TaxRate、DaysReceivable/Inventory/Payable、CapExRatio、DepreciationYears。`font.color=0000FF`；3–5 个 scenario 驱动 `fill=FFFF00`。≥3 次使用的设 named range（StartingARR、TaxRate、OpeningCash、GrowthRate、GrossMargin）。

**Step 3 — P&L 全公式**

行：Revenue / COGS / Gross Profit / OpEx / EBITDA / D&A / EBIT / Interest / EBT / Tax / Net Income。示例行图：B3=Revenue…B15=NI——按实际布局替换。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/model.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/P&L/B3",
        "props": { "formula": "Assumptions!B5", "font.color": "008000" }
      },
      {
        "command": "set",
        "path": "/P&L/C3",
        "props": { "formula": "B3*(1+Assumptions!C6)" }
      },
      {
        "command": "set",
        "path": "/P&L/D3",
        "props": { "formula": "C3*(1+Assumptions!D6)" }
      },
      {
        "command": "set",
        "path": "/P&L/E3",
        "props": { "formula": "D3*(1+Assumptions!E6)" }
      },
      {
        "command": "set",
        "path": "/P&L/B4",
        "props": { "formula": "-B3*(1-Assumptions!B7)" }
      },
      {
        "command": "set",
        "path": "/P&L/B5",
        "props": { "formula": "B3+B4" }
      }
    ],
    "stop_on_error": true
  }
}
```

跨表绿字；同表默认黑；`$` 行 `numFmt=$#,##0;($#,##0);"-"`。

**Step 4 — Balance Sheet**

Assets = Cash+AR+Inventory+Net PP&E；Liab = AP+Debt；Equity = Opening+RE。
`BS.Cash` **不是独立 plug**——必须等于 `'Cash Flow'!B<ending>`。
RE(t)=RE(t-1)+NI(t)−Div；Y1 Historical RE 可用 BS 恒等式 live formula，蓝字 + comment。

**Step 5 — Cash Flow**

Operating = NI+D&A−ΔWC；Investing = −CapEx；Financing = ΔDebt−Div。
Y2+ OpeningCash = 上期 EndingCash（同表自链：`C17=B19`…）。Y1 OpeningCash 来自 Assumptions。

**Step 6 — Balance check + cash recon**

示例行图：BS B10/B15/B17/B18；CF B5/B19/B21——按布局替换。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/model.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/Balance Sheet/B18",
        "props": {
          "formula": "IF(ABS(B10-B15-B17)<0.01,\"OK\",\"IMBALANCED: \"&ROUND(B10-B15-B17,0))",
          "bold": true,
          "font.color": "000000"
        }
      },
      {
        "command": "set",
        "path": "/Cash Flow/B21",
        "props": {
          "formula": "IF(ABS(B19-'Balance Sheet'!B5)<0.01,\"OK\",\"CF != BS CASH: \"&ROUND(B19-'Balance Sheet'!B5,0))",
          "bold": true
        }
      },
      {
        "command": "set",
        "path": "/Summary/B2",
        "props": {
          "formula": "'P&L'!E3",
          "font.color": "008000",
          "numFmt": "$#,##0;($#,##0);\"-\""
        }
      }
    ],
    "stop_on_error": true
  }
}
```

复制到 C/D/E。IMBALANCED 可用 CF `containsText`。

**Step 7 — cache refresh + format**

重设 summary / check / 跨表引用；列宽（A=28，B:E=15）；header fills 覆盖 A:E。

**Step 8 — Summary KPIs + ≥3 charts（board）**

最少 4 KPI：Revenue 27E、EBITDA Margin、Ending Cash、NI CAGR。Charts：Revenue&EBITDA column；margin line（先在 Summary 预置 ratio 行）；ending cash area。验证：BS B18:E18 / CF B21:E21 全 OK；Summary KPIs 非 null。

### Recipe B — DCF

Sheets：Assumptions、FCF（10yr）、WACC、DCF、Sensitivity。输出 Implied Equity Value + Per-Share + WACC×g 网格。

**Build order：** Assumptions → FCF → WACC → DCF → Sensitivity。

**Step 1 named ranges**

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/model.xlsx",
    "operations": [
      {
        "command": "add",
        "parent": "/",
        "type": "namedrange",
        "props": { "name": "WACC", "ref": "WACC!$B$12" }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "namedrange",
        "props": { "name": "TaxRate", "ref": "Assumptions!$B$8" }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "namedrange",
        "props": { "name": "TerminalGrowth", "ref": "Assumptions!$B$15" }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "namedrange",
        "props": { "name": "NetDebt", "ref": "Assumptions!$B$20" }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "namedrange",
        "props": { "name": "SharesOut", "ref": "Assumptions!$B$21" }
      }
    ],
    "stop_on_error": true
  }
}
```

**Step 2–3** FCF 十年全公式；WACC 面板（Rf/ERP/Beta/Re/Rd/weights/WACC）——输入蓝、派生黑。

**Step 4 — TV + NPV + equity bridge**

Notes 列用 `value`，永不 `formula` 散文（否则 `#NAME?`）。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/model.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/DCF/B3",
        "props": { "value": "Terminal value (Gordon growth)" }
      },
      {
        "command": "set",
        "path": "/DCF/C3",
        "props": {
          "formula": "FCF!K11*(1+TerminalGrowth)/(WACC-TerminalGrowth)",
          "font.color": "008000",
          "numberformat": "$#,##0;($#,##0);\"-\""
        }
      },
      {
        "command": "set",
        "path": "/DCF/C4",
        "props": {
          "formula": "SUMPRODUCT(FCF!B11:K11/(1+WACC)^FCF!B2:K2)",
          "font.color": "008000",
          "numberformat": "$#,##0;($#,##0);\"-\""
        }
      },
      {
        "command": "set",
        "path": "/DCF/C6",
        "props": {
          "formula": "C4+C5",
          "bold": true,
          "numberformat": "$#,##0;($#,##0);\"-\""
        }
      },
      {
        "command": "set",
        "path": "/DCF/C8",
        "props": {
          "formula": "C6+C7",
          "bold": true,
          "numberformat": "$#,##0;($#,##0);\"-\""
        }
      },
      {
        "command": "set",
        "path": "/DCF/C9",
        "props": {
          "formula": "C8/SharesOut",
          "bold": true,
          "numberformat": "$0.00"
        }
      }
    ],
    "stop_on_error": true
  }
}
```

`NPV` 与 `SUMPRODUCT` 均可缓存；不规则日期用 `XNPV`。补齐 C5=PV terminal、C7=−NetDebt。

**Step 5 — 5×5 WACC × g sensitivity**

每格独立公式替换网格轴；不要 Excel Data Table。先建 FCF/WACC/DCF，再单独 Batch 网格。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/model.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/Sensitivity/C15",
        "props": { "value": 0.075, "numFmt": "0.0%", "font.color": "0000FF" }
      },
      {
        "command": "set",
        "path": "/Sensitivity/D14",
        "props": { "value": 0.015, "numFmt": "0.0%", "font.color": "0000FF" }
      },
      {
        "command": "set",
        "path": "/Sensitivity/D15",
        "props": {
          "formula": "(NPV($C15,FCF!$B$11:$K$11)+(FCF!$K$11*(1+D$14)/($C15-D$14))/(1+$C15)^10-NetDebt)/SharesOut",
          "numFmt": "$0.00"
        }
      },
      {
        "command": "add",
        "parent": "/Sensitivity",
        "type": "conditionalformatting",
        "props": {
          "type": "colorscale",
          "sqref": "D15:H19",
          "minColor": "FFCDD2",
          "midColor": "FFFFFF",
          "maxColor": "C8E6C9"
        }
      }
    ],
    "stop_on_error": true
  }
}
```

复制公式填满网格时引用轴单元格，不可回指全局 `WACC` named range。Batch 后 Run `get` 抽查中心格，并检查 `#REF!` / `#DIV/0!`。

### Recipe C — LBO

Sheets：Assumptions、S&U、Debt、P&L、CF、Exit、Returns。输出 MOIC、IRR、可选 waterfall。

**Build order：** Assumptions → S&U → P&L → Debt → CF → Exit → Returns。写环前启用 iterate。

**Step 1 Sources & Uses**

Uses = Purchase_EV + fees + refinanced。Sources = Senior+Mezz+Revolver+Sponsor。
Sponsor 二选一：**stated** 或 **solved**——永不同时硬编码 stated 又 plug。S&U 平衡检查；stated≠plug → Gate 4 拒绝。

**Step 2 Debt schedule**

每 tranche 每年：Begin / Mandatory amort / Sweep / End / Avg / Interest。Senior：1% amort + excess sweep；Mezz：interest-only。Revolver：`MIN(capacity, …)` 外层封顶。Sweep comment 用独立 `add type comment`（不是 cell prop）。

**Step 3 P&L interest from Debt → circular**

Interest → NI → CF → Sweep → Debt → Interest。写环前：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/model.xlsx",
      "/",
      "--prop", "calc.iterate=true",
      "--prop", "calc.iterateCount=100",
      "--prop", "calc.iterateDelta=0.001"
    ]
  }
}
```

`#REF!` 或发散 = 停；修代数，不要抬 `iterateCount` 到 1000。

**Write-order surgery（复杂环）：**
1. De-ring：环细胞先写字面 `0`。
2. 写下游非环链。
3. Re-ring：close residents 后逐格 Run `set` 真公式。
Acceptance：环格 cachedValue 非零非 null。仍死锁则留 `=0` + comment `"circular; recalculates in Excel on F9"`，交付时披露。

**Step 4–5 CF + Exit/Returns**

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/model.xlsx",
    "operations": [
      {
        "command": "set",
        "path": "/Exit/B3",
        "props": {
          "formula": "'P&L'!F8*Assumptions!$B$25",
          "numberformat": "$#,##0;($#,##0);\"-\""
        }
      },
      {
        "command": "set",
        "path": "/Returns/B3",
        "props": {
          "formula": "'Exit'!B5/('S&U'!B9)",
          "numberformat": "0.00\"x\""
        }
      },
      {
        "command": "set",
        "path": "/Returns/B4",
        "props": {
          "formula": "IRR({-'S&U'!B9,0,0,0,0,'Exit'!B5})",
          "numberformat": "0.0%"
        }
      }
    ],
    "stop_on_error": true
  }
}
```

Labels：`comment` 元素（不是 cell prop）；Notes 列用 `value`。中期分红用 `XIRR`。验证：S&U BALANCED；MOIC/IRR 合理；`query cell:contains("#REF!")` 为 0。

## Sensitivity & scenarios

三模式择一：
(a) Base/Upside/Downside 列；(b) 下拉 + INDEX/MATCH；(c) 2 轴网格。
混 (a)+(b) 会造成循环输入。网格每格引用**轴单元格**。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/model.xlsx",
    "operations": [
      {
        "command": "add",
        "parent": "/Summary",
        "type": "validation",
        "props": {
          "sqref": "B1",
          "type": "list",
          "formula1": "Base,Upside,Downside"
        }
      },
      {
        "command": "set",
        "path": "/Assumptions/B5",
        "props": {
          "formula": "INDEX(C5:E5,MATCH(Summary!$B$1,$C$4:$E$4,0))"
        }
      },
      {
        "command": "add",
        "parent": "/Assumptions",
        "type": "comment",
        "props": {
          "ref": "B5",
          "text": "Revenue growth — picked by Summary!B1 scenario dropdown"
        }
      }
    ],
    "stop_on_error": true
  }
}
```

Football-field：Low/High/width stacked bar（invisible Low + visible width）。

## Financial function patterns

| Prefer | Over | Why |
|---|---|---|
| XNPV | NPV | 不规则日期 |
| XIRR | IRR | 不规则日期 / 多符号变化 |
| INDEX/MATCH | VLOOKUP | 插列安全 |
| IFERROR 或 IF(y=0) | 裸除法 | 防 `#DIV/0!` |
| MIRR | 多符号 IRR | 再投资假设 |
| SUMIFS | 复杂 SUMPRODUCT 数组 | 意图更清晰 |

规则：写你意思的公式，再 readback cachedValue。仅当 `#OCLI_NOTEVAL!` 持续存在才 fallback：先重设；仍不行 → 蓝字硬编码 + comment + 交付披露。

## Circular references & iterative calc

仅在代数正当环启用 iterate。验证收敛：读环格 → 轻推假设再推回 → 值应回到一致。`#REF!` / 发散 → 修代数，不要抬 iterateCount。

## Audit & Delivery Gate

**假定存在问题。**

### Gates 1–3 继承 xlsx

view issues、error cells（含 `#REF!`）、validate。

### Gate 4 — statement integrity

`IMBALANCED` / `CF !=` / `S&U IMBALANCE` 计数必须 0。常见因：跨表 `\!`——用 Batch JSON 重写。

### Gate 5 — cached-value on valuation cells

DCF C4/C5/C6/C8/C9；LBO Exit/Returns；3-stmt Summary KPIs。null / `#OCLI_NOTEVAL!` / `#REF!` → 拒绝。

### Gate 6 — hardcode / zone + named-range audit

Calc 零硬编码 Number；named ranges ≥3 且每个被 ≥1 formula 引用（查 `formula~=`，不要用 `:contains` 扫显示值）。

### Gate 5b — visual

`view html`：无 `###`、无截断、无 placeholder、checks 显示 OK、charts y=0、sensitivity 色阶、无 stale 0 KPI。

### Gate 6.1 — token sweep

`view text` 无 TBD / lorem / xxxx / `{{`。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": ["/workspace/model.xlsx", "/DCF/C8"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": ["/workspace/model.xlsx", "html"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": ["/workspace/model.xlsx"]
  }
}
```

可选：`save` flush 后再 `close`。

### Honest limit

`validate` 过 schema ≠ finance 正确。硬编码 BS.Cash 强行平衡、NPV cache 0、网格先于 FCF、未引号 `P&L` → `#NAME?`——靠 Gate 4/5/6/5b。

## Known Issues & Pitfalls

基座跨表 `!` / dotted props / renderer → xlsx。Model-specific：

- **AP 符号与负 COGS。** 若 P&L 将 COGS 存为负数，Accounts Payable 公式必须取反：`=-COGS*DaysPayable/365`；若 COGS 为正数则不取反。符号错误会放大 NWC 并翻转 cash-flow 方向，而且 `validate` 仍可能静默通过。
- `#NAME?` 未引号 sheet 名（Gate 5b）
- iterate 静默未收敛
- resident + 环 batch 死锁 → write-order surgery
- html 预览 stale 0 → cache-refresh
- sensitivity 建序错误
- BS.Cash 必须链接 CF；Y2+ OpeningCash 自链
- waterfall total 色约定；SharesOut 公式需蓝假设锚定

超出 Help 的 XML：`raw-set`/`add-part` 仅 prose + **approval**；**relationship**/path 依赖逐步 Run；**永不**进 Batch JSON。

## Help pointer

疑问时用 `{{OFFICE_HELP_TOOL}}`（format `xlsx` + topic）。Help 是权威 schema；本 Skill 是 financial-model 决策层。
