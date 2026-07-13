---
name: officecli-docx
description: "任何涉及 .docx 文件的场景都使用此 Skill，包括创建 Word 文档、报告、信函、备忘录或提案；读取、解析或提取 .docx 文本；编辑、修改或更新既有文档；处理模板、修订、批注、页眉/页脚或目录。当用户提到 'Word doc'、'document'、'report'、'letter'、'memo' 或 .docx 文件名时触发。"
---

# OfficeCLI DOCX Skill

## ⚠️ Help 优先规则

**本 Skill 说明高质量 docx 应达到的标准，而不是罗列全部命令 flag。遇到不确定的 property 名称、enum 值或 alias 时，先查 Help，切勿猜测。**

```bash
officecli help docx                         # List all docx elements
officecli help docx <element>               # Full element schema (e.g. paragraph, field, numbering, watermark, toc)
officecli help docx <verb> <element>        # Verb-scoped (e.g. add field, set section)
officecli help docx <element> --json        # Machine-readable schema
```

Help 与已安装的 CLI 版本一致。如本 Skill 与 Help 不一致，**以 Help 为准**。

## 心智模型

`.docx` 是由 XML part（`document.xml`、`styles.xml`、`numbering.xml`、`header*.xml`、`footer*.xml`、`comments.xml` 等）组成的 ZIP。用户看到的 heading、table、页码、TOC、tracked change 都存储在其中。`officecli` 在其上提供语义 path API（如 `/body/p[1]/r[2]`），通常无需直接操作 raw XML；确有必要时使用 `raw-set`（见 XML 附录）。

## Shell 与执行规范

docx path 包含 `[]`，部分 prop value 包含 `$`；它们都是 shell 元字符。转义分为三层，务必区分。

1. **Shell。** element path 必须加引号：使用 `"/body/p[1]"`，不要使用 `/body/p[1]`（zsh/bash 会将 `[N]` 视为 glob）。任何含 `$` 的 value 都必须整体用单引号包裹，例如 `--prop text='$50M'`。未加引号的 `$50M` 会变成 `M`；长字符串中混用 `'…$var…'` 与 `"…$50…"` 也会让 `$50` 悄然丢失。
2. **CLI（`text=`）。** `--prop text=` 会解释双字符转义 `\n` 与 `\t`：前者变为 `<w:br/>` soft line break，后者变为 `<w:tab/>`。docx / pptx / xlsx 的行为一致。若要写入字面的反斜杠加 n，使用 `\\n`（通常不需要）。这同样适用于 table 行级快捷方式 `c1…cN`（cell 内的 `\n` 也会变为 `<w:br/>`）。
3. **JSON（`batch`）。** 在 `batch` 的 JSON heredoc 中，字符串 `"\n"` 也可以传递真实的换行符，结果相同。

如果有疑问，请在写入后 `view text` 并逐个字符进行比较。

**增量执行。** `officecli` 每次调用都会修改文件。每次只运行一条命令，并检查其 exit code；含 50 条命令的脚本若在第 3 条失败，后续操作会静默连锁失败。完成任一结构操作（新增 style、table、TOC、section break）后，先运行 `get` 确认结果，再继续叠加操作。

**打开/保存生命周期：**开始时运行 `officecli open <file>`，结束时运行 `officecli save <file>` flush 到磁盘。`save` 只写入并保留 resident 以供后续编辑；仅在需要一次性交接时使用 `officecli close <file>` 释放 resident。两者始终安全，不会报错或丢失工作。对同一 style 的多个 paragraph，使用 `batch`（一次传递整个 array）。**只在非 officecli 边界 flush：**officecli 自身读取始终能看到编辑；仅在非 OfficeCLI 程序（python-docx、Word、renderer、交付流程）读取文件前运行 `save`/`close`。

**`$FILE` 约定。** 所有命令均使用 `"$FILE"`；只需设置一次（`FILE="your-doc.docx"`）。不要把示例中的字面量 `doc.docx` / `review.docx` 复制到产物中，必须替换为真实目标文件。

## 输出标准

每个 document 都必须满足以下交付标准；执行命令前先理解它们。

**清晰的层级。** 每个非简单 document 都应具备 Title → Heading 1 → Heading 2 → body，而非一堵未经 style 处理的 `Normal` paragraph 墙。若 `view outline` 仅显示扁平列表，说明层级缺失。

**显式设置 heading size**（Word 默认 style size 会因 template 而漂移）：**H1 ≥ 18pt**（长报告用 20pt）、H2 = 14pt bold、H3 = 12pt bold、body = 11–12pt、line spacing = 1.15–1.5x。优先使用 `style=Heading1` 而非 inline size，这样重新应用 theme 时只需修改一次定义；但无法信任 template 的 style 时，仍需显式 `set` size。

**一种 body font、一个 accent。** 使用一种可读的 body font（Calibri、Cambria、Georgia、Times New Roman）；标题可使用 accent color 或 table header 强调，不要使用彩虹式格式。

**通过 property 控制间距。** 在 paragraph 上使用 `spaceBefore` / `spaceAfter`。连续空 paragraph 会破坏分页，并被 `view issues` 标记。

**排版质量。**新增内容使用 curly quote（`'` `'` `"` `"`），不要使用 ASCII quote；可直接写 Unicode，或在 `raw-set` 中使用 XML entity（`&#x2018;`/`&#x2019;`/`&#x201C;`/`&#x201D;`）。范围使用 en-dash `–`（`2024–2026`），插入语使用 em-dash `—`。

**任何超过 1 页的 document 都应具有 header、footer 和页码。**页码必须使用 live `PAGE` field（`--prop field=page`），不要写入字面量 "Page 1"；CLI 会自动注入 `<w:fldChar>`（见“页眉与页脚”）。

**保留现有 template。** 编辑已有视觉样式的文件时，应遵循原有约定；它们优先于本指南。

### 视觉交付底线（适用于每个 document）

声明完成前，运行 `officecli view "$FILE" html` 并读取返回的 HTML path，确认以下全部条件：

- **不得将 placeholder token 渲染为数据。** `$xxx$`、`{var}`、`{{name}}`、`<TODO>`、`lorem`、`xxxx` 不得出现在 heading、body、cover、TOC、caption、header 或 footer。需要由人工填写的字面量 `{name}` 只能放在明确的说明 paragraph 中（如“发送前替换 `{name}`”），不能作为成品内容。
- **不得出现被截断的 title 或溢出的 cell。** 加宽 column 或设置 `wrapText`，不要靠裁剪内容解决。
- **document 有 3 个以上 heading 时必须包含 TOC**（`--type toc`）。
- **cover 至少填充 60%，最后一页至少填充 40%。** 封面不足时补充 subtitle / author / date / scope / key highlight；最后的“Thank you”页可补充 conclusion / next step / contact / legal。
- **document 文本中不得出现字面量 `\$`、`\t`、`\n`。** 若 `view text` 显示它们，说明 shell escape 泄漏；删除该 paragraph 后重新输入。

如果有任何失败，请在宣布完成之前停止并修复。

## 通用工作流程

六个步骤，适用于每个非简单构建。

1. **打开文件。** 运行 `officecli open "$FILE"`（默认启用 resident）。新文件先运行 `officecli create "$FILE"`。
2. **了解现状。** 对既有文件先运行 `officecli view "$FILE" outline`，查看 heading tree、section count，以及 TOC / watermark / tracked change 是否已存在。切勿盲目编辑。
3. **增量构建。** 顺序为：structure → content → formatting，即 style 与 numbering definition → section / page setup → heading / body → table / image / field / TOC → header / footer → comment。每次结构操作后都运行 `get` 核对，再继续。
4. **按规范格式化。** 显式 heading size、spacing、width、alignment、tab 和 list indent 都是交付内容，不是可选润色。
5. **保存，但以 structure 而非 cached text 为准。** `officecli save "$FILE"` 会写入 XML。TOC / PAGE / NUMPAGES / SEQ / PAGEREF field 的 **cached value** 在人工重新计算前（Word 中按 F9）可能为空或过期。应确认 field *存在*（`get --depth 3` 能找到 `<w:fldChar>`），不要只相信可见文本。
6. **QA：假定存在问题。** 最后一条命令 exit code 为 0 并不代表完成；经过一次完整的修复与验证循环且没有发现新问题，才算完成。见 QA。

## 快速入门

最小可行 docx：一个 heading、一个 body paragraph、一个 subheading，以及带 live page-number field 的 footer。请按自己的文件与内容调整，不要直接复制粘贴。

```bash
FILE="review.docx"
officecli create "$FILE"
officecli open "$FILE"
officecli add "$FILE" /body --type paragraph --prop text="Q4 2026 Review" --prop style=Heading1 --prop size=20pt --prop bold=true --prop spaceAfter=12pt
officecli add "$FILE" /body --type paragraph --prop text="Revenue grew 18% year-over-year, ahead of plan." --prop size=11pt --prop spaceAfter=8pt
officecli add "$FILE" /body --type paragraph --prop text="Key Drivers" --prop style=Heading2 --prop size=14pt --prop bold=true --prop spaceBefore=12pt --prop spaceAfter=6pt
officecli add "$FILE" /body --type paragraph --prop text="Enterprise renewals, upsell, and a new EMEA region." --prop size=11pt
officecli add "$FILE" / --type footer --prop type=default --prop size=9pt --prop text="Page " --prop field=page
officecli set "$FILE" "/footer[1]/p[1]" --prop align=center
officecli save "$FILE"
officecli validate "$FILE"
```

已验证：`validate` 返回 `no errors found`；`get /footer[1] --depth 3` 显示由 5 个 run 组成的 PAGE field chain（begin / instrText / separate / cached value / end）。

## 阅读与分析

先宽后窄。`outline` 展示已有内容；明确位置后，再使用 `view text` / `get` / `query` 深入查看。

```bash
officecli view "$FILE" outline            # heading tree, section count, table/image counts, watermark, tracked-changes presence — orient here first
officecli view "$FILE" html               # Read the returned HTML path: first visual check after a batch of edits (hierarchy, empty-para spacing, missing TOC)
officecli view "$FILE" text --start 1 --end 80   # text for content QA; paths shown as [/body/p[N]] so you can jump back with get
officecli view "$FILE" annotated          # values + style/font/size + warnings per run
officecli view "$FILE" stats              # paragraph counts, font usage, style distribution
officecli view "$FILE" issues             # empty paras, missing alt text, spacing anomalies
```

`officecli watch "$FILE"` 会持续运行 live preview，供人工用户按需打开；agent 自检使用 `view html`。最终的视觉验证应由用户在 Word / WPS / Pages 中打开 `.docx` 完成。

**检查单个 element。** 使用 XPath 风格的语义 path（索引从 1 开始）。务必加引号，避免 shell 将 `[N]` 解释为 glob。最后一个 element 应使用 `[last()]`（带括号），`[last]` 会报错。添加 `--json` 可获得 machine-readable output。

```bash
officecli get "$FILE" /                          # document root: metadata, page setup
officecli get "$FILE" "/body/p[1]"                # one paragraph
officecli get "$FILE" "/body/p[1]/r[1]"           # one run (character-level formatting)
officecli get "$FILE" "/body/tbl[1]" --depth 3    # table with rows and cells
officecli get "$FILE" "/footer[1]" --depth 3      # footer — check for fldChar
officecli get "$FILE" "/styles/Heading1"          # style definition
officecli get "$FILE" /numbering --depth 2        # numbering abstractNum + num bindings
```

**跨 document 查询。** 使用 CSS-like selector 做系统性检查，而非手动遍历。operator 包括：`=`、`!=`、`~=`（包含）、`>=`、`<=`、`[attr]`（存在）。完整说明见 `officecli query --help`。

```bash
officecli query "$FILE" 'paragraph[style=Heading1]'       # all H1s
officecli query "$FILE" 'p:contains("quarterly")'         # text match
officecli query "$FILE" 'p:empty'                         # empty paragraphs (clutter)
officecli query "$FILE" 'image:no-alt'                    # accessibility gaps
officecli query "$FILE" 'paragraph[size>=24pt]'           # numeric comparison
officecli query "$FILE" 'field[fieldType!=page]'          # fields other than PAGE
```

`query --json` 会将结果包装在 `.data.results[]` 中；使用 `jq '.data.results | length'` 计数。

**大型 document。** 用 `view outline` 按 heading 导航，再用 `query` 跳转；不要把整个 body 全部载入上下文。

## 创建和编辑

可用 verb：`add`（新增 element）、`set`（修改 prop）、`remove`、`move`、`swap`、`batch`、`raw-set`（最后手段的 XML）。一次构建中，90% 都是 paragraph、run、table、少量 image、一个 TOC 和一个 footer。

### Paragraph、run 与 style

paragraph（`p`）是一个块；run（`r`）是其中 character formatting 一致的一段 span。在 `p` 上设置 paragraph-level prop（style、alignment、spacing、indent）；在 `r` 上设置 font / size / color / bold。

```bash
officecli add "$FILE" /body --type paragraph --prop text="Executive Summary" --prop style=Heading1 --prop size=18pt --prop bold=true --prop spaceAfter=12pt
officecli set "$FILE" "/body/p[1]/r[1]" --prop color=1F4E79
```

垂直间距使用 `spaceBefore` / `spaceAfter`，不要堆叠空 paragraph。左缩进使用 `--prop indent=720`（twips）；首行缩进使用 `firstLineIndent=360`，悬挂缩进使用 `hangingIndent=720`。前导空格会触发 `view issues`。

### 表格

table 位于 `/body/tbl[N]`，由 row `tr[N]` 与 cell `tc[N]` 组成。先按 row/column count 创建，再填充内容。

```bash
officecli add "$FILE" /body --type table --prop rows=4 --prop cols=3 --prop width=100%
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop header=true --prop c1=Quarter --prop c2="Revenue" --prop c3="Growth"
officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[1]/r[1]" --prop bold=true
```

行级 `set` 支持 `height`、`header` 与 `c1 / c2 / … / cN` text shortcut（`cN` 可适用于任意 column count）。cell 格式（bold、fill、color）必须设置在 cell 内的 paragraph / run，**不能**设在 row 上。对单个 cell 的 border，在 `tc` 上设置 cell-level `border.*`（例如 `--prop border.bottom="single;6;000000;0"`），或在内部 paragraph 上设置 paragraph-level `pbdr.*`。

**水平分隔线应使用 paragraph bottom border，不要使用单行 table。** table 作为分隔线会渲染出带最小高度的空白框，尤其会破坏 header/footer。请在 paragraph 上使用 `pbdr.bottom`（`STYLE;SIZE;COLOR`）：

```bash
officecli set "$FILE" "/body/p[3]" --prop pbdr.bottom="single;6;2E75B6"
```

### 列表（项目符号、编号、多级）

单级 bullet/number 应在 paragraph 上设置 `listStyle`（`listStyle` 是 paragraph prop，**不是** run prop，这是常见错误）：

```bash
officecli add "$FILE" /body --type paragraph --prop text="First item" --prop listStyle=bullet
```

多级编号（法律文书风格 1 / 1.1 / 1.1.1）依次 `add` `abstractNum` 与 `num`，再让每个 paragraph 引用 `numId`：

```bash
officecli add "$FILE" /numbering --type abstractnum --prop format=decimal     # → abstractNum id=0
officecli add "$FILE" /numbering --type num --prop abstractNumId=0             # → num id=1
officecli add "$FILE" /body --type paragraph --prop text="Section one" --prop numId=1 --prop ilvl=0
```

ID 从 0 开始：第一个 `abstractNum` 的 id 为 0；`num` 通过 `abstractNumId=0` 引用它，并会获得 id=1。不存在的 `abstractNumId` 会报错，因此创建后应检查 id。使用 `officecli query "$FILE" 'paragraph[numId>0]'` 验证。level 与 format 选项见 `help docx abstractnum` / `help docx num`。

### Tab stop（签名行、leader row）

tab stop 是 paragraph 的 first-class `tab` child；`pos` 支持 `6in` / `6cm` / twips，`val` 的取值为 `left` / `center` / `right`，`leader` 的取值为 `none` / `dot` / `hyphen` / `underscore`。见 `help docx tab`。

```bash
officecli add "$FILE" "/body/p[1]" --type tab --prop pos=6in --prop val=right --prop leader=dot
```

**Leader 注意事项。** 单独设置 `leader=dot` 不会生成点线；只有文本与 tab stop 之间的 run 中包含真实 `<w:tab/>` 字符时才会显示。做法是先定义 stop（`add tab --prop pos=6in --prop val=right --prop leader=dot`），再将 `\t` 写入文本：`--prop text="Chapter 1\t12"`。`\t` 会变为 `<w:tab/>`，点线会延伸至右对齐页码。（字面量 `text="Chapter 1 ......... 12"` 也能交付，但真实 tab stop 对齐更稳定。）

### Field（PAGE / NUMPAGES / DATE / MERGEFIELD / REF）

field 是在 render 时计算的 live value。`fieldType` 指定 field；`name` 给出目标（merge name 或 `ref` bookmark）；`format` / `instr` 添加 switch。

| Field | 用途 | 示例 |
|---|---|---|
| `page` | 当前页码 | 在 footer 上使用 `--prop field=page`，或 inline 使用 `--prop fieldType=page` |
| `numpages` | 总页数 | `--prop field=numpages` / `--prop fieldType=numpages` |
| `date` | 当天日期 | `--prop fieldType=date --prop format='yyyy-MM-dd'` |
| `mergefield` | template merge token | `--prop fieldType=mergefield --prop name=CustomerName` |
| `ref` | 对 bookmark 的 cross-reference | `--prop fieldType=ref --prop name=bookmarkName` |

完整的 `fieldType` enum（30+ 个值，包括 `pageref`、`seq`、`styleref`、`docproperty`、`createdate` 等）见 `help docx field`。**不存在 `fieldInstr` fieldType**；当 typed shortcut 不足以表达时，使用 `instruction` prop 写入 raw field instruction text。picture switch（`MERGEFIELD Amount \# "#,##0.00"`、`DATE \@ "yyyy年MM月"`）应通过 `--prop instruction='…'` 传入；`mergefield` 的 `format` prop 会被忽略并产生 warning，应使用 `instruction`。

```bash
officecli add "$FILE" "/body/p[3]" --type field --prop fieldType=mergefield --prop name=customer_name
# Renders «customer_name» — visible placeholder, replaced in Word at mail-merge time.
```

**MERGEFIELD template：绝不能渲染字面量 placeholder。** 出现在 body 中的 `{{customer_name}}` 或 `$NAME$` 表示 template 失败，收件人会直接看到它。应插入真实 MERGEFIELD（见上方），或仅将字面 token 放入清晰的说明 paragraph。使用 `query 'field[fieldType=mergefield]'` 确认。

**SEQ / PAGEREF / TOC field value。** officecli 写入时不保存已 render 的 field value；请按 path 的实际需求重新计算：

- **SEQ 编号**（`Figure 1/2/3`）：`officecli set "$FILE" / --prop recalcFields=seq` 会按 body document order 统计 SEQ field，并写入 cached value（`evaluated` 会变为 true；switch/format 见 `help docx document`）。相对 heading 的 `\s` 以及 header/footer 内的 SEQ 仍由 Word 处理。
- **PAGE / PAGEREF / NUMPAGES / TOC page number** 依赖分页，而 officecli 没有 pagination engine；运行 `officecli set "$FILE" /settings --prop updateFields=true`，让 Word 在打开时计算它们。

多图 document 应同时使用两种方式。学术论文见 `officecli-academic-paper` Skill。

### Header 与 footer（页码）

单命令模式：CLI 会注入 `<w:fldChar>`，无需手写 field：

```bash
# Empty first-page footer — auto-enables differentFirstPage so the cover has no page number
officecli add "$FILE" / --type footer --prop type=first --prop text=""
# Default footer with live page number
officecli add "$FILE" / --type footer --prop type=default --prop align=center --prop size=9pt --prop text="Page " --prop field=page
```

两者都存在时，default footer 是 `/footer[2]`；只有它时则是 `/footer[1]`。**验证：**`get --depth 3` 必须显示 `fldChar` child，而不只是带字面量 `"Page"` 的 run（`view outline` 会为 live field 和 static text 都输出 "Footer: Page"，不可据此判断）。不要执行 `set --prop differentFirstPage=true`：该 prop 不受支持，会以 exit 2 拒绝，不会静默失败。添加 first-type footer 会自动设置该 bit。复合 **Page X of Y** 见 recipe (b)。

### 目录

任何有 3+ 个 heading 的 document：

```bash
# Built-in Heading styles are TOC sources.
officecli add "$FILE" /body --type paragraph --prop text="Introduction" --prop style=Heading1
# Custom paragraph styles need an outline level (0 = Heading 1).
officecli add "$FILE" /styles --type style --prop id=ThesisH1 --prop type=paragraph --prop outlineLvl=0
# Add the TOC only after its sources exist.
officecli add "$FILE" /body --type toc --prop levels="1-3" --prop title="Table of Contents" --prop hyperlinks=true --index 0
```

page number 默认会 render（可用 `--prop pageNumbers=true` 显式开启）。可直接访问 TOC：`/toc[1]` 或 `/tableofcontents` 会解析为第一个 TOC field，供 `get` / `set` / `remove` 使用，无需手动遍历 XPath。

**TOC 交付步骤（交接前必须执行）。** live TOC field 在重新计算前只是 placeholder。部分 viewer 会在首次打开时填充它，另一些则会一直显示字面量 `Update field to see table of contents`，直至读者重新计算。按收件人的能力选择：

- **会重新计算（或按 F9）：**运行 `officecli set "$FILE" /settings --prop updateFields=true`，使 Word 在打开时计算 TOC（以及所有 field）；也可添加可见提示“按 F9 刷新 TOC 和页码”。
- **不能或不会重新计算：**使用**静态 TOC fallback — recipe (f)**。live field 在重新计算前只能显示 placeholder，任何 headless pipeline 都无法提前填充它。

交付检查：只要读者不会重新计算，`officecli query "$FILE" 'p:contains("Update field to see")'` 就必须返回空；若有匹配，切换至 recipe (f)。

### 图片

picture 位于 run 内。为满足 accessibility 要求，创建时必须直接传入 alt text：

```bash
officecli add "$FILE" "/body/p[5]" --type picture --prop src=logo.png --prop width=1.5in --prop alt="Acme logo"
```

交付前确认 `officecli query "$FILE" 'image:no-alt'` 结果为空。

### 图表

展示数据时，使用**原生 chart**：可编辑、可应用 theme、具备 accessibility，并可在 Word 中重新 render；绝不能用 chart 的平面 PNG screenshot 代替。每个 series 使用 `data="Label:v1,v2,…"`；每个 series 对应一个 `data=`（或使用 `series1=` / `series2=`）。

```bash
officecli add "$FILE" /body --type chart --prop chartType=bar --prop title="Revenue by Region" --prop categories="EMEA,APAC,Americas" --prop data="2026:120,150,180"
```

`chartType` 可为 `bar` / `column` / `line` / `pie` / `area` / `scatter`（axis、legend、series style 见 `help docx chart`）。只有 officecli 无法构建的特殊 chart 才可退回为 `--type picture` 的 PNG。

### 超链接和 bookmarks

外部链接通过 `hyperlink`：

```bash
officecli add "$FILE" "/body/p[2]" --type hyperlink --prop url="https://example.com" --prop text="our site"
```

**内部链接**（到 bookmark）使用 `--prop anchor=bookmarkName` — 而不是 `url` 中的 `#fragment`：

```bash
officecli add "$FILE" "/body/p[2]" --type hyperlink --prop anchor=chapter1 --prop text="See Chapter 1"
```

另一种选择是将 `PAGEREF` field 与可见文本配对。请参阅 `help docx hyperlink` / `help docx bookmark`。

### 节与页面设置

document root `/` 包含 page setup（`pageWidth`、`pageHeight`、margin，单位为 twips）。多 section document（如横向插页、column）需要 `add` `section` break，见 `help docx section`。支持 camelCase（规范写法，如 `pageWidth`）与 lowercase alias（如 `pagewidth`），优先使用 camelCase。

```bash
officecli set "$FILE" / --prop pageWidth=12240 --prop pageHeight=15840 --prop marginTop=1440 --prop marginLeft=1440
# Newspaper-style multi-column flow (columnSpace in twips; 720 = 0.5in):
officecli set "$FILE" / --prop columns=2 --prop columnSpace=720
```

### 强制分页：双重保障

有两种机制，但**任一机制单独使用都不能在所有 viewer 中可靠生效**。不同 viewer 或前置内容可能忽略 `<w:pageBreakBefore/>`，也可能把 `<w:br w:type="page"/>` 渲染为 soft break。对需要另起页的每个 H1、TOC heading 和 cover 末尾 paragraph，同时应用两者：

```bash
# Default: apply the break directly to the heading.
officecli add "$FILE" /body --type paragraph --prop text="Introduction" --prop style=Heading1 --prop pageBreakBefore=true
```

`--prop break=newPage` 是 `pageBreakBefore=true` 的简写 alias（接受 `newPage|page|nextPage|pageBreak`）。两者生成相同 XML，仍应遵循上述双重保障规则。使用 `view html` 预览并核对页数。

### 报告级 recipe

以下模式常见于长篇报告，且均已实际执行并通过 `validate`。

**(a) 内容丰富的 cover：达到 ≥ 60% 填充底线。** 依次加入保密 banner、title、subtitle、客户/项目/日期块和 key-theme strip，再将下一 section 强制移至新页：

```bash
officecli add "$FILE" /body --type paragraph --prop text="CONFIDENTIAL — CLIENT USE ONLY" --prop align=center --prop size=9pt --prop color=C00000 --prop spaceAfter=24pt
officecli add "$FILE" /body --type paragraph --prop text="Strategic Growth Review" --prop style=Title --prop size=32pt --prop bold=true --prop align=center --prop font=Cambria --prop spaceAfter=8pt
officecli add "$FILE" /body --type paragraph --prop text="FY26 Outlook and Scenario Planning" --prop italic=true --prop size=16pt --prop align=center --prop spaceAfter=36pt
officecli add "$FILE" /body --type paragraph --prop text='Prepared for: Acme Corp. Leadership Team' --prop align=center --prop size=11pt
officecli add "$FILE" /body --type paragraph --prop text='Engagement: 2026-04 — 2026-06' --prop align=center --prop size=11pt --prop spaceAfter=36pt
officecli add "$FILE" /body --type paragraph --prop text="Key themes: 1) margin resilience, 2) EMEA expansion, 3) capital allocation." --prop align=center --prop italic=true --prop size=10pt
officecli add "$FILE" /body --type paragraph --prop text="Executive Summary" --prop style=Heading1 --prop pageBreakBefore=true
```

**(b) Page X of Y footer：组合 PAGE + NUMPAGES。** 先创建 footer paragraph，再通过三个 child operation 构成 live `Page <X> of <Y>`。这是 `help docx footer` 的官方 recipe。

```bash
officecli add "$FILE" / --type footer --prop type=default --prop text="Page " --prop align=center --prop size=9pt
officecli add "$FILE" "/footer[1]/p[1]" --type field --prop fieldType=page
officecli add "$FILE" "/footer[1]/p[1]" --type run --prop text=" of "
officecli add "$FILE" "/footer[1]/p[1]" --type field --prop fieldType=numpages
officecli get "$FILE" "/footer[1]/p[1]" --depth 1 | grep -o fldChar | wc -l   # expect ≥ 4; use grep -o ... | wc -l, NOT grep -c (single-line XML returns 1)
```

**(c) Header row 使用填充色与白色 bold text。** 顺序很重要：先填入 header cell text（空 cell 中没有 run；对空 cell 执行 `set …/tc[N]/p[1]/r[1]` 会报“未找到”），再设置 cell fill，最后设置 run format：

```bash
officecli add "$FILE" /body --type table --prop rows=5 --prop cols=4 --prop width=100%
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop header=true --prop c1=Quarter --prop c2=Revenue --prop c3=Growth --prop c4=Status
for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[$col]" --prop fill=1F4E79
  officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[$col]/p[1]/r[1]" --prop bold=true --prop color=FFFFFF
done
for row in 3 5; do for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[$row]/tc[$col]" --prop fill=D9E2F3      # zebra stripe
done; done
```

**(d) Financial table：数值右对齐、total bold、total row 的 bottom border。**

```bash
for row in 2 3 4 5; do for col in 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[$row]/tc[$col]/p[1]" --prop align=right
done; done
for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[5]/tc[$col]/p[1]/r[1]" --prop bold=true
  officecli set "$FILE" "/body/tbl[1]/tr[4]/tc[$col]/p[1]" --prop pbdr.bottom="single;6;000000;0"
done
```

**(e) 含多个 bullet 的 cell（SWOT / risk matrix）。** `c1="a\nb"` 会在**同一个** paragraph 中生成 `<w:br/>` 换行，适合纯多行文本；bullet 则需要独立 paragraph。先用 `set c1=` 写入首项，再在同一 cell 下对每个后续 bullet `add paragraph`（使用 `listStyle=bullet`）：

```bash
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop c1="Installed base of 18k enterprise seats"
officecli add "$FILE" "/body/tbl[1]/tr[1]/tc[1]" --type paragraph --prop text="Margin structure above peer median" --prop listStyle=bullet
officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[1]" --prop listStyle=bullet
```

若首个 paragraph 落在底部，使用以下命令重新排序：`officecli move "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[N]" --index 0`。

**(f) Static TOC fallback（跨 viewer 稳定）。** 交付给不会自动重新计算的 viewer 时，live TOC field 会显示字面量 `Update field to see table of contents`。纯 CLI pipeline 无法像 Word 一样在保存时预填 TOC field。解决方法是 `remove` TOC field，保留可见 heading，并为每个 heading 手动写入带 dot leader 的目录项。

```bash
officecli query "$FILE" 'p:contains("Update field to see")'        # note the /body/p[N] paths, then:
officecli remove "$FILE" "/body/p[N]"                              # repeat per hit
officecli add "$FILE" /body --type paragraph --prop text="Contents" --prop style=TOCHeading --prop size=14pt --prop bold=true --index <pos>
officecli add "$FILE" /body --type paragraph --prop text="1. Executive Summary ......................................... 3" --prop size=11pt --index <pos+1>
# … one per heading. Page numbers manual; eyeball positions via view html. Live --type toc remains correct for recipients who recalculate.
```

### Template 交付：分离 Template Notes 与最终用户内容

HR / legal / vendor template 常含仅供内部使用的说明（如“替换 `{{CompanyName}}`”），不得随成品交付。可使用两种方式：

- **末尾的 “Template Notes” section：**放在明确的 `Heading 1`（“Template Notes for HR Users”）下，所有说明置于其后；分发前，从该 heading 起向下 `remove`（先用 `query 'paragraph[style=Heading1]:contains("Template Notes")'` 定位）。
- **由 bookmark 限定的内部 section：**放在 `__template_notes_start` / `_end` bookmark 之间；交付时通过 `raw-set` 移除两锚点之间的全部内容。

template 的交付 gate：移除后，`query 'p:contains("Template Notes")'` 和 `query 'p:contains("{{")'` 均必须返回空。若 notes paragraph 仍存在，下游员工会看到内部说明。

### 高级主题（撰写普通报告时可跳过）

报告、memo、letter、proposal 和 HR template 不需要本节。仅当 document 属于 academic（equation、footnote、bibliography）、reviewed（comment、tracked change）或 marked（watermark）时继续阅读。

**Equation 与 footnote。** `--type equation` 接受 LaTeX，支持 `\frac`、`\sum`、Greek letter、`\mathit`、`\mathcal` 等。默认会创建独立的 `/body/oMathPara[N]` display block；若需 inline `<m:oMath>`，请在 paragraph parent path 上使用 `--prop mode=inline`（`add "/body/p[N]" --type equation --prop formula=… --prop mode=inline`）。footnote 按 paragraph index 自动编号。bibliography 的 hanging indent：每个条目使用 `firstLineIndent=-720 indent=720`。

```bash
officecli add "$FILE" /body --type equation --prop formula="\\frac{a}{b} + \\sum_{i=1}^{n} x_i"
officecli add "$FILE" "/body/p[3]" --type footnote --prop text="See Appendix A for methodology."
```

**Comment 与 tracked change。** 批量 accept/reject 使用 `set "$FILE" /revision --prop revision.action=accept`（或 `--prop revision.action=reject`）；可用 selector 缩小范围，如 `/revision[@author=Alice]` 或 `/revision[@type=ins]`。使用 `query ins` 与 `query del` 查找单项变更（`trackedchange` 不是 selector）。在 run 上通过 `--prop revision.type=ins|del --prop revision.author=…` 创建 tracked change；完整 `revision.*` set（含 `format` / `moveFrom` / `moveTo`）见 `help docx run`。添加 comment：`add "/body/p[4]" --type comment --prop author=… --prop text=…`；以 `--prop parentId=N` 建立 reply thread，以 `set "/comments/comment[N]" --prop done=true` 标记为 resolved。应 resolve 而非 delete，以保留 audit trail；`query 'comment[done=false]'` 可列出未解决项。prop schema 见 `help docx comment` / `help docx run`。

**Watermark。** 使用一条命令 `add / --type watermark --prop text="DRAFT" --prop color=BFBFBF --prop opacity=0.8` 添加（默认 opacity 为 0.5）；之后可通过 `set /watermark --prop opacity=…` 调整。

**何时切换 Skill。** chapter draft、≤ 3 个 footnote、≤ 2 个 equation、没有 bibliography/cross-reference 时继续使用 docx。涉及 citation style（APA / Chicago / IEEE / GB 7714）、in-text ↔ reference 自动链接、含 `\ref` 的编号 equation、“List of Figures”或自动更新 cross-reference 时，切换至 **`academic-paper`**。document 的目标是**数据采集**时，切换至 **`officecli-word-form`**，如可填写 form、带用户填写槽的 contract、questionnaire、mail-merge template（`<w:sdt>` content control、`<w:ffData>`、`documentProtection=forms`）。

### raw-set escape hatch（L1 / L2 / L3）

分为三级精度；始终选用能完成任务的最低级别。

- **L1 — high-level prop**（`--prop text=…`、`--prop style=Heading1`）：默认方式，覆盖 80% 场景。
- **L2 — dotted-attribute fallback**（`pbdr.top=`、`ind.left=`、`shd.fill=`、`padding.top=`、`font.size=`）：L1 缺少所需设置时使用。例如 `--prop pbdr.bottom="single;6;1F4E79;0"`；会生成 schema-valid XML。
- **L3 — `raw-set` with XML**：最后手段，没有 schema protection。用于 internal hyperlink、composite field 等 typed verb 无法表达的结构（见 XML 附录）。

border 格式为 `style;size;color;space`，例如 `single;4;FF0000;1`。hex color 不带 `#`，即 `FF0000`。scheme color name（`accent1..6`、`dark1` / `dark2`、`light1` / `light2`、`hyperlink`）可用于任意接受 hex color 的位置；如需跨 theme 保持稳定，优先使用 hex color。

## QA（必须执行）

**假定存在问题：QA 是找 bug，不是确认步骤。** 第一次生成的 document 几乎不会完全正确；首次检查发现零问题，通常表示检查不够仔细。heading 看似正常，直到 `view outline` 显示 H1 下直接出现 H3；footer 看似显示 “Page 1”，直到 `get --depth 3` 发现它其实是 static run 而非 field。

### 声明“完成”前的最小循环

1. `officecli view "$FILE" issues`：检查 empty paragraph、missing alt text、formatting anomaly。
2. `officecli view "$FILE" outline`：检查 heading hierarchy（不得从 H1 跳到 H3）、TOC 是否存在、section count。
3. `officecli view "$FILE" text --max-lines 400`：检查 typo、残留的 `\$` / `\t` / `\n` 字面量与 placeholder token。
4. `officecli validate "$FILE"`：执行 schema check（Delivery Gate 会在关闭且写入磁盘后的文件上再次执行）。
5. **视觉检查：将整个 document 作为 contact sheet。** 仅适用于具备视觉能力的 agent；若无法解读图像，跳过此步，并在交接时标记 document 为“未做视觉验证”。运行 `officecli view "$FILE" screenshot --grid auto -o /tmp/sheet.png`，再读取图片。`--grid auto` 会将**每一页**平铺在一张图像中（自动选择 column count；`--grid 4` 可强制指定），可直观看到 pagination、blank page、heading rhythm、失衡的 margin 与 TOC/cover 位置，而不只是 DOM。Windows+Word 通过真实 Word render 每页，其他平台使用 HTML。若 screenshot 失败，退回 `view html`，并将跨页 break / alignment / rhythm 标记为“未做视觉验证”。thumbnail 只用于**定位**；可疑页应使用无 `--grid` 的 `screenshot --page N` 以 full resolution 确认细节（column alignment、line spacing、indent、dark-on-dark、caption placement；Windows 使用真实 Word）。`validate` 通过不等于可交付；应达到“看起来像真实 document”的质量。
6. 发现任何问题后先修复，再**重新运行完整循环**；一个修复常会引入另一个问题。

### Delivery Gate（交付前运行；任一失败即 REJECT）

复制以下内容，设置 `FILE`。所有 gate 打印 OK 前，不能声明完成。

```bash
FILE="your-file.docx"

# Gate 1 — schema.
officecli close "$FILE" 2>/dev/null
officecli validate "$FILE" | grep -q "no errors found" || { echo "REJECT Gate 1: validate failed"; exit 1; }
echo "Gate 1 OK"

# Gate 2 — token leak (shell-escape / template tokens / literal \$ \t \n). grep -c never false-PASSes.
LEAK=$(officecli view "$FILE" text | grep -cE '(\$[A-Za-z_]+\$|\{\{[^}]+\}\}|<TODO>|xxxx|lorem|\\[\$tn])')
[ "$LEAK" -eq 0 ] && echo "Gate 2 OK" || { echo "REJECT Gate 2: $LEAK leak line(s)"; officecli view "$FILE" text | grep -nE '(\$[A-Za-z_]+\$|\{\{[^}]+\}\}|<TODO>|xxxx|lorem|\\[\$tn])'; exit 1; }
# A TOC placeholder is valid before a Word-compatible field engine updates it; confirm the TOC field and updateFields setting structurally instead.

# Gate 3 — live PAGE field exists when a footer is expected.
FLD=$(officecli query "$FILE" 'field[fieldType=page]' --json | jq '.data.results | length')
[ "$FLD" -ge 1 ] && echo "Gate 3 OK" || { echo "REJECT Gate 3: no live PAGE field"; exit 1; }
echo "Delivery Gate PASS"
```

### Field / cached-value spot-check

field 的 cached value 在写入时可能过期或为空；应通过**structure 而不是 text**确认其存在。

- **Footer PAGE：**`get /footer[N] --depth 3` 会列出 begin / instrText / separate / cached / end 的 run chain：单一 PAGE 至少 5 个 run，组合 “Page X of Y” 至少 11 个。只有一个包含 `"Page"` 的 run 表示 field 缺失；使用 `--prop field=page` 重新添加。
- **TOC：**`get /toc[1] --depth 2` 显示 field structure。在重新计算前，page number 可能显示为 `1 1 1 1` 或 `Update field to see…`（见 TOC 交付步骤）。
- **MERGEFIELD：**`query 'field[fieldType=mergefield]'` 应为每个 slot 返回一个结果，其他位置不应存在字面量 `{{name}}`。

### 已知局限

`validate` 只能发现 schema error，不能发现 design error。heading hierarchy 错误、伪 Heading 1 size、placeholder token 被作为 body text、没有 cover 的 document 却带空 first-page footer，都可能通过 `validate`。contact-sheet visual check（`screenshot --grid`）与 field structure 检查用于发现这些问题。

### QA 显示说明（无需处理）

- `view text` 会对每个 numbered list item 都显示 `"1."`，无论实际 render 的数字是多少；实际输出会正确递增。
- `view issues` 可能在 cover paragraph、居中 title、list item、bibliography entry 上标记“body paragraph 缺少首行缩进”。只有 APA/academic body text 需要首行缩进；在 block-style professional document 中这是预期行为。

## 已知问题与陷阱

内容“看起来损坏”时，先归因：**[AGENT-ERROR]** 表示 document 有误，应修复；**[RENDERER-BUG]** 表示 document 正确但 viewer render 不同，不应继续追逐；**[skill gap]** 表示本 Skill 未覆盖相应规则，应提出问题。

### Renderer quirk（跨 viewer，[RENDERER-BUG]，无需处理）

判定 color / field / chart 有问题前，先在用户的目标 viewer 中打开文件；若目标 viewer 中正常，则这是 viewer quirk。

- **PAGE field 可能显示字面量 “Page”**（没有数字），直到重新计算；应以 `fldChar` 是否存在判断，而不是以数字判断。
- **TOC 的 cached page number 可能显示为 “1 1 1 1”**，直到按 F9。
- **部分 viewer 中 pie / doughnut fill 可能塌缩为一种颜色**（`column` / `bar` render 正常）。
- **form-control checkbox 可能渲染为双框**；**OMML equation baseline** 也可能因 viewer 而异（XML 相同）。

### 常见陷阱

| 陷阱 | 正确做法 |
|---|---|
| `--index` 与 `[N]` | `--index` 从 0 开始；`[N]` path 从 1 开始 |
| 多次 `add --index N` 使用相同 N | 每次插入都会将后续内容下移；复用 N 会让后插入项位于先插入项之前。应逆序插入，或以 `paraId` 为锚点使用 `move --after` / `--before` |
| zsh/bash 中未加引号的 `[N]` | 每个 path 都加引号：`"/body/p[1]"` |
| 将 `[last]` 用作 predicate | 必须使用带括号的 `[last()]` |
| spacing 中使用原始 twips | 使用带单位的 value：`12pt`、`0.5cm`、`1.5x` |
| 用空 paragraph 制造 spacing | 使用 `spaceBefore` / `spaceAfter` |
| 对 row 执行 cell format 的 `set` | row `set` 只支持 `height`、`header`、`c1..cN` text；format 应设在 cell paragraph / run |
| 在 run 上设置 `listStyle` | `listStyle` 是 paragraph property |
| 依赖前导空格缩进 | 使用 `indent=720` / `firstLineIndent=360` / `hangingIndent=720`（dotted `ind.left` / `ind.firstLine` 也可） |
| 用 `set differentFirstPage=true` 隐藏 cover page number | 不支持；添加 first-type footer：`--type footer --prop type=first --prop text=""` |
| 单独使用 `--type pagebreak` 或 `pageBreakBefore` | 两者同时使用（见“强制分页”） |
| 把多个 bullet paragraph 合并到一个 cell | `c1="a\nb"` 会得到 `<w:br/>` 换行（同一 paragraph）；独立 bullet paragraph 使用 recipe (e) |
| 何时用 dotted property / `raw-set` | 优先选择 L2 dotted property，L3 `raw-set` 仅作最后手段 |
| 后一个 paragraph 继承了前一个 heading style | 在下一个 paragraph 显式设置 `--prop style=Normal` |
| 修改仍在 Word 中打开的文件 | 先关闭 Word |
| batch heredoc 在 `$` / `'` 处中断 | 使用单引号 delimiter 的 heredoc：`cat <<'EOF' \| officecli batch …` |

## raw-set XML 附录（L3 模式）

`raw-set` 直接注入 literal OOXML，没有 schema protection。`<w:pPr>` 中 element 的顺序为：`pStyle`、`numPr`、`spacing`、`ind`、`jc`、`rPr`（最后）。curly quote 使用 XML entity（`&#x2018;` / `&#x2019;` / `&#x201C;` / `&#x201D;`）。任何带前导或尾随空格的 `<w:t>` 都要添加 `xml:space="preserve"`。RSID 是 8 位 hex number。除非用户另有指定，tracked change 的 author/comment 使用 “Claude”。

**Tracked change 的插入/删除。** 优先在 run 上使用 high-level `--prop revision.type=ins|del`；仅当 typed path 无法表达时才使用 `raw-set`（例如拒绝/恢复其他 author 的变更）。应替换完整的 `<w:r>…</w:r>`，不要把 tag 注入 run 内；复制原 `<w:rPr>` 到两个 run 以保留 format。在 `<w:del>` 内使用 `<w:delText>`（field instruction 则使用 `<w:delInstrText>`）：

```xml
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30</w:delText></w:r></w:del>
<w:ins w:id="2" w:author="Claude" w:date="2026-01-01T00:00:00Z"><w:r><w:t>60</w:t></w:r></w:ins>
<w:r><w:t> days.</w:t></w:r>
```

删除 paragraph / list item 的所有内容时，还要在 paragraph mark（`<w:pPr><w:rPr>`）内写入 `<w:del/>`；否则 accept change 后会留下空 paragraph。要**拒绝另一 author 的插入**，将自己的 `<w:del>` 嵌套在对方的 `<w:ins>` 中；要**恢复对方的删除**，在后方添加 `<w:ins>`，不要修改对方的原始结构。

**Internal hyperlink 到 bookmark**（优先使用 high-level `--prop anchor=` path；`raw-set` 仅用于命令无法表达的自定义 run style）：

```xml
<w:hyperlink w:anchor="chapter1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>See Chapter 1</w:t></w:r></w:hyperlink>
```

**将 field 组合在一个 run 中**（例如单条 command path 无法组合两个 field 时）：使用 `fldChar begin / instrText / separate / value / end` chain：

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

**Comment marker** 与 `<w:r>` 是 sibling，不能位于其中。reply thread 与 resolved status 属于 high-level 功能，可使用 `--prop parentId=` / `done=`（见上方 comment 说明）：

```xml
<w:commentRangeStart w:id="0"/><w:r><w:t>annotated text</w:t></w:r><w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
```

使用 `officecli set "$FILE" /settings --prop updateFields=true` 强制在打开时重新计算 field（写入 `<w:updateFields w:val="true"/>`，适用于依赖 layout 的 PAGE / PAGEREF / NUMPAGES / TOC page number，不需要 `raw-set`）。对 SEQ 编号，优先使用 `set / --prop recalcFields=seq`；它会直接写入正确的 cached value，无需等待 Word。

### Help 速查

遇到疑问时使用：`officecli help docx`、`officecli help docx <element>`、`officecli help docx <verb> <element>`；`--json` 用于 agent。Help 是权威 schema，本 Skill 用于帮助决策。
