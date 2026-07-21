# Academic Paper Agent Skill

Author: xiesq, 2026-07-15

**Do not invent** citations, data, results, or venue requirements. Use only user-supplied facts and mark gaps explicitly.

**本 Skill 是 `docx`/`word` 之上的 scene layer。** docx 硬规则（style architecture、heading hierarchy、page-break、live PAGE field、Delivery Gate、renderer quirks）继承不重教。本文件只加学术增量：citation styles、equations、SEQ/PAGEREF、multi-column、bibliography hanging indent、abstract/keywords/affiliation。

docx 基座已覆盖处写 `→ see docx`。未读过基座则先 `{{OFFICE_LOAD_SKILL_TOOL}}` name=`word`。

## ⚠️ Help 优先规则

**本 Skill 说明学术论文需要什么，而不是罗列全部 flag。** 不确定的 prop / field instruction，先查 Help。

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "docx"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "docx",
      "equation"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "docx",
      "field"
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
      "word"
    ]
  }
}
```

`{{OFFICE_LOAD_SKILL_TOOL}}` 只读取指导，不安装或修改本地 skills。

## 心智模型与继承

### What "academic" means（identity）

六个 delta：

1. **Citation style 是契约。** APA / Chicago / IEEE / MLA 决定作者格式、日期位置、参考文献顺序、in-text 形态。
2. **Equations 是一等公民** — inline `oMath` 与 display `oMathPara`。
3. **Figures/tables 自动编号** — `SEQ Figure` / `SEQ Table`；`PAGEREF` 做 live 页码交叉引用。
4. **Bibliography 用 hanging indent**（不是 first-line indent）。
5. **Abstract / keywords / affiliation** 是首页三件套，block-style，无装饰。
6. **Multi-column** 出现在 IEEE/ACM/Nature 等：单栏 abstract + 双栏 body。

### Reverse handoff

白皮书 / policy brief / tech report / HR template（无 venue/citation）→ 留在 docx 基座。仅当至少具备两项 {citation biblio, equations, SEQ/PAGEREF, multi-column, abstract+keywords} 时用本 Skill。

## 宿主工具与执行规范

（对应原 CLI「Shell & Execution Discipline」；shell 引号改为 Tool JSON 传参。）

**Path 与特殊字符。** path 含 `[N]` 原样传入；DOI / `$` / 货币写在 JSON 字符串内即可。props 中的 `\n` / `\t` 仍由 Office 工具解释。

**增量执行。** 结构操作后检查再继续。节奏：一条（或一块独立 Batch）→ 检查 output → 继续。生命周期用 `create` / `open` / `save` / `close`。

**阶段 0 Batch/Run 契约（固定补丁，非业务加戏）。**

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。

## Workflow — 5 verbs

1. **Read venue spec.** APA 7 / Chicago 17 / IEEE / MLA 9 / journal-specific。
2. **Plan sections.** Abstract → keywords → intro → methods → results → discussion → conclusion → references。≥3 Heading1 → TOC（→ docx TOC）。
3. **Set styles up front.** Heading1/2/3、Caption、AbstractTitle、Bibliography——内容前定义。
4. **Build body in order.** Title block → abstract → keywords → TOC → body → figures/tables with SEQ → bibliography → footnotes last。
5. **QA.** 继承 docx Gates 1–3，加 academic Gates 4–5。

## Requirements（academic floor）

继承 docx Requirements。另加：

### Typography / spacing（venue-aware）

- Font：Times New Roman 11–12pt（默认）或 venue 指定（IEEE Times 10pt 2-col；APA 可 Calibri 11）。
- H1=20pt bold，H2=14pt bold，H3=12pt bold italic，body 11–12pt。
- Line spacing：APA 2x；Chicago/IEEE 常 1.5x；不低于 1.15x。
- Margins：1 inch（1440 twips）除非 venue 另有规定。

### Abstract / biblio / captions

- Abstract：**无** `firstLineIndent`；`spaceAfter=12pt`。`view issues` 对 Abstract 报 missing first-line indent 为假阳性——忽略。
- Bibliography：`indent=720 hangingIndent=720`（hanging 0.5"）。
- Figure caption **下方**；Table caption **上方**（APA/Chicago/IEEE/MLA 一致）。
- Citation round-trip：每个 in-text key 必须解析到参考文献（Gate 4）。
- 编号图/表必须有 live SEQ（Gate 5）——禁止硬编码 "Figure 1" 文本。

### Cover / first-page

Title（居中 20–22pt bold）、author(s)、affiliation、submission/journal、date、abstract、keywords。docx「cover ≥60% filled」仍适用。

### Section numbering（STYLE-DEPENDENT — 勿盲套）

| Style | H1 | H2 | Example |
|---|---|---|---|
| APA 7 | 居中 bold **无编号** | flush-left bold | `Introduction` |
| Chicago | `N. Title` | `N.M Title` | `1. Introduction` |
| IEEE | `N. TITLE` Roman ALL CAPS | `A. Subtitle` | `I. INTRODUCTION` |
| MLA 9 | 无编号 left bold | same | `Literature Review` |

APA 7 的 L1 必须居中 bold，L2 flush-left bold，L3 flush-left bold italic，L4/L5 为 run-in heading；APA heading 禁止添加 `1.` / `2.` 前缀。IEEE 的一级标题使用 Roman numeral + ALL CAPS，二级标题使用 `A.` / `B.` / `C.` + title case。Arabic body numbering 只属于 Chicago，不得套到 APA、IEEE 或 MLA。

References / Bibliography / Works Cited / Acknowledgments **一律无编号**。

## Quick Start — minimal APA

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": ["/workspace/paper.docx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "open",
    "command_arguments": ["/workspace/paper.docx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/paper.docx",
    "operations": [
      {
        "command": "set",
        "path": "/",
        "props": { "defaultFont": "Times New Roman" }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Remote Work and Team Cohesion",
          "align": "center",
          "size": "20pt",
          "bold": true,
          "spaceAfter": "24pt"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": { "text": "Alice Chen", "align": "center", "size": "12pt" }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Department of Psychology, Stanford University",
          "align": "center",
          "size": "11pt",
          "spaceAfter": "24pt"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Abstract",
          "align": "center",
          "size": "14pt",
          "bold": true,
          "spaceBefore": "12pt",
          "spaceAfter": "6pt"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "This study examines remote-work adoption on team cohesion across 18 months...",
          "size": "12pt",
          "lineSpacing": "2x",
          "spaceAfter": "12pt"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Keywords: remote work, team cohesion, psychological safety",
          "italic": true,
          "size": "11pt",
          "spaceAfter": "18pt"
        }
      }
    ],
    "stop_on_error": true
  }
}
```

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/paper.docx",
    "operations": [
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "1. Introduction",
          "style": "Heading1",
          "size": "20pt",
          "bold": true,
          "spaceBefore": "18pt",
          "spaceAfter": "12pt"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Remote-work research (Smith, 2024) has expanded since 2020...",
          "size": "12pt",
          "lineSpacing": "2x",
          "firstLineIndent": 720
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "References",
          "style": "Heading1",
          "size": "20pt",
          "bold": true,
          "spaceBefore": "18pt",
          "spaceAfter": "12pt"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Smith, J. (2024). Remote work and cohesion. Journal of Applied Psychology, 109(3), 412-430.",
          "size": "12pt",
          "lineSpacing": "2x",
          "indent": 720,
          "hangingIndent": 720
        }
      },
      {
        "command": "add",
        "parent": "/",
        "type": "footer",
        "props": { "type": "default", "align": "center", "size": "10pt", "field": "page" }
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
    "command_arguments": ["/workspace/paper.docx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": ["/workspace/paper.docx"]
  }
}
```

骨架通过 `validate`；真实论文再追加 body、hanging-indent 条目、带 SEQ 的图/表、≥3 H1 时 TOC。APA 正式稿若要求 L1 无编号居中，勿保留 Quick Start 中的 `1. Introduction` 前缀。

## Citation style recipes

| Style | In-text | Ref order | Spacing | Footnotes? |
|---|---|---|---|---|
| APA 7 | `(Smith, 2024)` | alpha | 2x | rare |
| Chicago Notes-Bib | superscript | alpha | 1.5–2x | **primary** |
| IEEE | `[1]` first-use order | citation order | 1.15–1.5x 2-col | rare |
| MLA 9 | `(Smith 412)` | alpha Works Cited | 2x | rare |

共享：hanging indent `indent=720 hangingIndent=720`；≥3 H1 加 TOC；`updateFields=true` 时 TOC 页码可能未计算。

### APA 7

In-text `(Author, Year)`；三作者以上首次后 `et al.`；引用列表 alpha；文章标题 sentence case；期刊名 title case italic；DOI 用 https URL。双倍行距含 abstract/references；body `firstLineIndent=720`。

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/paper.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Remote work adoption accelerated during the pandemic (Kramer & Kramer, 2020).",
      "--prop",
      "size=12pt",
      "--prop",
      "lineSpacing=2x",
      "--prop",
      "firstLineIndent=720"
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
      "/workspace/paper.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Kramer, A., & Kramer, K. Z. (2020). The potential impact of the Covid-19 pandemic on occupational status. Journal of Vocational Behavior, 119, 103442.",
      "--prop",
      "size=12pt",
      "--prop",
      "lineSpacing=2x",
      "--prop",
      "indent=720",
      "--prop",
      "hangingIndent=720"
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
      "/workspace/paper.docx",
      "/body/p[last()]",
      "--type",
      "hyperlink",
      "--prop",
      "url=https://doi.org/10.1016/j.jvb.2020.103442",
      "--prop",
      "text=https://doi.org/10.1016/j.jvb.2020.103442"
    ]
  }
}
```

QA：Run `query paragraph[hangingIndent]`；零参考文献用 first-line indent。

### Chicago 17 Notes-Bibliography

首次脚注全文；随后 shortened；连续同页 `Ibid.`；非连续 shortened（**不用** op. cit.）。Bibliography alpha hanging。可拆 Primary/Secondary Heading2。Author-Date 变体接近 APA，仅标点不同：`(Smith 2024)`。

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/paper.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=The Ming dynasty's maritime prohibition shaped coastal trade for two centuries.",
      "--prop",
      "size=12pt",
      "--prop",
      "lineSpacing=1.5x",
      "--prop",
      "firstLineIndent=720"
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
      "/workspace/paper.docx",
      "/body/p[last()]",
      "--type",
      "footnote",
      "--prop",
      "text=Timothy Brook, The Troubled Empire: China in the Yuan and Ming Dynasties (Cambridge, MA: Harvard University Press, 2010), 142."
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
      "/workspace/paper.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Bibliography",
      "--prop",
      "style=Heading1",
      "--prop",
      "size=20pt",
      "--prop",
      "bold=true",
      "--prop",
      "spaceBefore=18pt"
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
      "/workspace/paper.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Brook, Timothy. The Troubled Empire: China in the Yuan and Ming Dynasties. Cambridge, MA: Harvard University Press, 2010.",
      "--prop",
      "size=12pt",
      "--prop",
      "indent=720",
      "--prop",
      "hangingIndent=720"
    ]
  }
}
```

QA：`query footnote` 计数 ≥ 正文脚注锚点数。

### IEEE

`[1]` 按首次出现编号；双栏 body；Abstract 单栏 10pt 约 200–250 词；H1 Roman ALL CAPS；tables Roman（`Table I`）；figures 仍 Arabic。body `firstLineIndent=288`（约 0.2"）。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/ieee.docx",
    "operations": [
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Attention-based anomaly detection has been applied to industrial sensor data [1], [2].",
          "size": "10pt",
          "lineSpacing": "1.15x"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "[1] A. Smith and B. Jones, \"Attention for anomaly detection,\" IEEE Trans. Neural Netw., vol. 35, no. 2, pp. 412-430, 2024.",
          "size": "10pt",
          "indent": 720,
          "hangingIndent": 720
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "[2] C. Lee, \"Time-series anomaly survey,\" in Proc. ICML, 2023, pp. 1200-1215.",
          "size": "10pt",
          "indent": 720,
          "hangingIndent": 720
        }
      }
    ],
    "stop_on_error": true
  }
}
```

QA：最高 `[N]` = 参考文献条数。

### MLA 9

Diff vs APA：in-text `(Author Page)` **无逗号**；节名 **Works Cited**；九核心元素 period 分隔；书名 italic、文章标题引号。段落设置同 APA（2x、hanging indent）。条目内容必须来自用户提供书目——**Do not invent**。

## Equations（OMML）

| Mode | Visual | Use |
|---|---|---|
| `display`（默认） | 独立居中块 `<m:oMathPara>` | 编号方程、定理 |
| `inline` | 跑在正文里 `<m:oMath>` | 散文中的变量 |

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/paper.docx",
    "operations": [
      {
        "command": "add",
        "parent": "/body",
        "type": "equation",
        "props": { "mode": "display", "formula": "x^2 + y^2 = z^2" }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "equation",
        "props": { "mode": "display", "formula": "\\lambda_1 + \\alpha" }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "equation",
        "props": {
          "mode": "display",
          "formula": "\\frac{1}{2\\pi} \\int_0^{\\infty} e^{-x^2} dx"
        }
      }
    ],
    "stop_on_error": true
  }
}
```

Inline：先 add paragraph 前缀文字，再对 `p[last()]` Run add `equation mode=inline` + trailing `run`。验证：`view text` 出现 Unicode 数学；`raw /document` 可见 `<m:oMathPara`。正文变量/希腊字母/下标必须走 `equation mode=inline`，禁止写进 `paragraph text=`。

**LaTeX pitfalls：** `\left(...\right)` / `\left[...\right]` 内含上下标可能 parse error；外层 `^2` 可以。用 plain `()` / `[]`，display 下 OMML 会自适应。`move` `/body/oMathPara[N]` 重排包装段；`before` 可能留空段，优先 `index` / `after`。表格 cell 路径不能直接挂 equation——用 `tc[N]/p[1]` + `mode=inline`。

**编号方程：** 无原生 `\eqno`。同一段用 center+right tab：`[tab] equation(inline) [tab] (1)`。单栏默认 text width≈8300 twips → center 4150、right 8300；双栏按 gutter 重算。勿把编号拆到下一行。Schema：Help `docx tab`。

## Figures / tables / SEQ + PAGEREF

原生 fieldType：`seq`、`pageref`。全部 caption 加完后一次：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/paper.docx",
      "/",
      "--prop", "recalcFields=seq"
    ]
  }
}
```

Figure caption BELOW image：

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/paper.docx",
      "/body",
      "--type",
      "picture",
      "--prop",
      "src=arch.png",
      "--prop",
      "width=5in"
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
      "/workspace/paper.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Figure ",
      "--prop",
      "style=Caption",
      "--prop",
      "size=10pt",
      "--prop",
      "italic=true",
      "--prop",
      "align=center"
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
      "/workspace/paper.docx",
      "/body/p[last()]",
      "--type",
      "field",
      "--prop",
      "fieldType=seq",
      "--prop",
      "identifier=Figure"
    ]
  }
}
```

再 Run 追加 caption run（`: Attention-based...`）+ `bookmark name=fig_arch`。交叉引用：paragraph 文字 + `fieldType=pageref name=fig_arch`。Table：caption ABOVE，再 `add table`。

预 recalc 的 SEQ 在 `view text` 显示 `#OCLI_NOTEVAL!{SEQ Figure}`；recalc 后应为升序数字。

## Footnotes vs endnotes

Footnote = 页底；Endnote = 文末。`view annotated` 里 reference run 可能显示空串——用 `query footnote` 或 `get /footnotes/footnote[N]` 确认。脚注不移动 paragraph 索引；body 完成后再加。

## Bibliography section

节名随风格：References（APA/IEEE/Chicago Author-Date）/ Bibliography（Chicago Notes-Bib）/ Works Cited（MLA）。每条独立 paragraph + hanging indent。canonical pair：`indent=720 hangingIndent=720`（非 `ind.firstLine=-720`）。Round-trip：in-text markers vs hanging entries（Gate 4）。

## Multi-column（IEEE）

单栏 title/authors/abstract → `section type=continuous` → 对 post-break section `columns=2` → body →（可选）再 section + `columns=1` 回退。**回退不是可选**——否则参考文献也双栏。每次 section break 插入空 paragraph，索引 +1；用 `get /body depth 1` 重索引。可用 `/section[last()]`。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/ieee.docx",
    "operations": [
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Attention-Based Anomaly Detection for Industrial Time Series",
          "align": "center",
          "size": "18pt",
          "bold": true,
          "spaceAfter": "12pt"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "paragraph",
        "props": {
          "text": "Abstract",
          "align": "center",
          "size": "12pt",
          "bold": true,
          "spaceAfter": "6pt"
        }
      },
      {
        "command": "add",
        "parent": "/body",
        "type": "section",
        "props": { "type": "continuous" }
      }
    ],
    "stop_on_error": true
  }
}
```

再 Run `set /section[2]` → `columns=2`、`columnSpace=1cm`；再 Batch IEEE body（`I. INTRODUCTION` 等）。视觉：`view html` 确认 abstract 全宽、body 双栏。

## Abstract / keywords / affiliation

Title → authors（可 superscript 多机构）→ affiliations → date → Abstract heading → block abstract（无 firstLineIndent，150–300 词）→ keywords italic。cover ≥60% filled。多机构：`run superscript=true`。Running header：`header type=default`；封面跳过见 docx headers。`evenAndOddHeaders` + `header type=even` 见 Help。

## QA — Delivery Gate

**假定存在问题。**

### Gates 1–3 继承 docx

schema validate、token leak、live PAGE structure。

### Gate 4 — citation round-trip

数 in-text markers vs hanging-indent entries；citations ≤ entries（IEEE 用 `[N]`；APA/MLA 调正则）。

### Gate 5a — SEQ presence + distinct numbers

有 Figure/Table 可见标签则必须有 SEQ；`recalcFields=seq` 后 distinct 升序。用 `query field[fieldType=seq]` 计数（勿仅靠 raw grep）。

### Gate 5b — Visual audit（强制）

`view html` + Read：页序、Abstract 仅一次、Fig/Table 编号不重复、方程是数学不是 `lambda_1` 纯文本、IEEE Roman caps、APA 无编号居中 H1、交叉引用可解析、H1/H2/H3 视觉可区分。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": ["/workspace/paper.docx", "/body", "--depth", "1"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": ["/workspace/paper.docx", "html"]
  }
}
```

可选人工预览：`watch` 或在 Word / WPS / Pages 打开。

### Honest limit

`validate` 过 schema ≠ 学术正确。APA 文用 IEEE 引用、禁止脚注的风格里塞脚注、硬编码 Figure 1 漂移——靠 Gate 4/5。

## Known Issues（academic-specific）

基座 pitfalls → docx。Academic：

- `\left`+内嵌上下标 parse error；`move` oMathPara 留空段
- section break +1 paragraph offset；`/section[last()]` 可用
- multi-column 不自动回退
- equation on `tc[N]` 被拒（改 `tc[N]/p[1]` inline）
- SEQ 靠 `recalcFields=seq`，非 per-field raw 修补
- hanging indent 规范形式；footnote 空 run
- caption 放置方向：`validate` 不抓

超出 Help 的 XML 修补：`raw-set`/`add-part` 仅 prose + **approval**；依赖 **relationship**/path 则逐步 Run；**永不**进 Batch JSON。

## Help pointer

疑问时用 `{{OFFICE_HELP_TOOL}}`（format `docx` + topic）。Help 是权威 schema；本 Skill 是学术决策层。
