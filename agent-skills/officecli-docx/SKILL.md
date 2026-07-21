# DOCX Agent Skill

Author: xiesq, 2026-07-15

## ⚠️ Help 优先规则

**本 Skill 说明高质量 docx 应达到的标准，而不是罗列全部命令 flag。遇到不确定的 property 名称、enum 值或 alias 时，先查 Help，切勿猜测。**

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
      "paragraph"
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
      "add",
      "field"
    ]
  }
}
```

Help 与已安装的 Office 工具版本一致。如本 Skill 与 Help 不一致，**以 Help 为准**。

`{{OFFICE_LOAD_SKILL_TOOL}}` 只读取指导，不安装或修改本地 skills。

## 心智模型

`.docx` 是由 XML part（`document.xml`、`styles.xml`、`numbering.xml`、`header*.xml`、`footer*.xml`、`comments.xml` 等）组成的 ZIP。用户看到的 heading、table、页码、TOC、tracked change 都存储在其中。宿主工具在其上提供语义 path API（如 `/body/p[1]/r[2]`），通常无需直接操作 raw XML；确有必要时使用 `raw-set`（见 XML 附录；仅 prose + 宿主 **approval**，**永不**进 Batch JSON）。

## 宿主工具与执行规范

（对应原 CLI「Shell 与执行规范」；shell 引号改为 Tool JSON 传参。）

**Path 与特殊字符。** docx path 含 `[]`，部分 prop value 含 `$`：

- path 原样传入 `command_arguments` / Batch `path`（含 `[N]`），无需 shell 引号。
- `$` 写在 JSON 字符串内即可（无 shell history expansion）。
- props 中的 `\n` / `\t` 仍由 Office 工具解释：`\n` 变为 `<w:br/>` soft line break，`\t` 变为 `<w:tab/>`（docx / pptx / xlsx 一致）；字面量反斜杠+n 写 `\\n`。这同样适用于 table 行级快捷方式 `c1…cN`。
- Batch JSON 中字符串 `"\n"` 也可传递真实换行符，结果相同。

如果有疑问，请在写入后 Run `view text` 并逐个字符进行比较。

**增量执行。** 每次结构操作后检查结果再继续。含 50 条操作的流程若第 3 条失败，后续会连锁失败。正确节奏：一条（或一块独立 Batch）→ 检查 output → 继续。完成任一结构操作（新增 style、table、TOC、section break）后，先 Run `get` 确认结果，再继续叠加操作。

**阶段 0 Batch/Run 契约（固定补丁，非业务加戏）。**

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。

**打开/保存生命周期：**开始时 Run `open`，结束时 Run `save` flush 到磁盘。`save` 只写入并保留 resident 以供后续编辑；仅在需要一次性交接时 Run `close` 释放 resident。两者始终安全，不会报错或丢失工作。对同一 style 的多个 paragraph，使用 Batch（一次传递整个 array）。**只在非 Office 工具边界 flush：**宿主自身读取始终能看到编辑；仅在非 Office 工具之外的程序（python-docx、Word、renderer、交付流程）读取文件前运行 `save`/`close`。

示例 path 使用 `/workspace/doc.docx`；必须替换为真实目标文件。

## 输出标准

每个 document 都必须满足以下交付标准；执行命令前先理解它们。

**清晰的层级。** 每个非简单 document 都应具备 Title → Heading 1 → Heading 2 → body，而非一堵未经 style 处理的 `Normal` paragraph 墙。若 `view outline` 仅显示扁平列表，说明层级缺失。

**显式设置 heading size**（Word 默认 style size 会因 template 而漂移）：**H1 ≥ 18pt**（长报告用 20pt）、H2 = 14pt bold、H3 = 12pt bold、body = 11–12pt、line spacing = 1.15–1.5x。优先使用 `style=Heading1` 而非 inline size，这样重新应用 theme 时只需修改一次定义；但无法信任 template 的 style 时，仍需显式 `set` size。

**一种 body font、一个 accent。** 使用一种可读的 body font（Calibri、Cambria、Georgia、Times New Roman）；标题可使用 accent color 或 table header 强调，不要使用彩虹式格式。

**通过 property 控制间距。** 在 paragraph 上使用 `spaceBefore` / `spaceAfter`。连续空 paragraph 会破坏分页，并被 `view issues` 标记。

**排版质量。**新增内容使用 curly quote（`'` `'` `"` `"`），不要使用 ASCII quote；可直接写 Unicode，或在 `raw-set` 中使用 XML entity（`&#x2018;`/`&#x2019;`/`&#x201C;`/`&#x201D;`）。范围使用 en-dash `–`（`2024–2026`），插入语使用 em-dash `—`。

**任何超过 1 页的 document 都应具有 header、footer 和页码。**页码必须使用 live `PAGE` field（`field=page`），不要写入字面量 "Page 1"；CLI 会自动注入 `<w:fldChar>`（见“页眉与页脚”）。

**保留现有 template。** 编辑已有视觉样式的文件时，应遵循原有约定；它们优先于本指南。

### 视觉交付底线（适用于每个 document）

声明完成前，Run `view html` 并读取返回的 HTML path，确认以下全部条件：

- **不得将 placeholder token 渲染为数据。** `$xxx$`、`{var}`、`{{name}}`、`<TODO>`、`lorem`、`xxxx` 不得出现在 heading、body、cover、TOC、caption、header 或 footer。需要由人工填写的字面量 `{name}` 只能放在明确的说明 paragraph 中（如“发送前替换 `{name}`”），不能作为成品内容。
- **不得出现被截断的 title 或溢出的 cell。** 加宽 column 或设置 `wrapText`，不要靠裁剪内容解决。
- **document 有 3 个以上 heading 时必须包含 TOC**（`type=toc`）。
- **cover 至少填充 60%，最后一页至少填充 40%。** 封面不足时补充 subtitle / author / date / scope / key highlight；最后的“Thank you”页可补充 conclusion / next step / contact / legal。
- **document 文本中不得出现字面量 `\$`、`\t`、`\n`。** 若 `view text` 显示它们，说明 escape 泄漏；删除该 paragraph 后重新输入。

如果有任何失败，请在宣布完成之前停止并修复。

## 通用工作流程

六个步骤，适用于每个非简单构建。

1. **打开文件。** Run `open`（默认启用 resident）。新文件先 Run `create`。
2. **了解现状。** 对既有文件先 Run `view outline`，查看 heading tree、section count，以及 TOC / watermark / tracked change 是否已存在。切勿盲目编辑。
3. **增量构建。** 顺序为：structure → content → formatting，即 style 与 numbering definition → section / page setup → heading / body → table / image / field / TOC → header / footer → comment。每次结构操作后都 Run `get` 核对，再继续。
4. **按规范格式化。** 显式 heading size、spacing、width、alignment、tab 和 list indent 都是交付内容，不是可选润色。
5. **保存，但以 structure 而非 cached text 为准。** `save` 会写入 XML。TOC / PAGE / NUMPAGES / SEQ / PAGEREF field 的 **cached value** 在人工重新计算前（Word 中按 F9）可能为空或过期。应确认 field *存在*（`get depth 3` 能找到 `<w:fldChar>`），不要只相信可见文本。
6. **QA：假定存在问题。** 最后一条命令成功并不代表完成；经过一次完整的修复与验证循环且没有发现新问题，才算完成。见 QA。

## 快速入门

最小可行 docx：一个 heading、一个 body paragraph、一个 subheading，以及带 live page-number field 的 footer。请按自己的文件与内容调整，不要直接复制粘贴。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": [
      "/workspace/review.docx"
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
      "/workspace/review.docx"
    ]
  }
}
```

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/review.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Q4 2026 Review",
      "--prop",
      "style=Heading1",
      "--prop",
      "size=20pt",
      "--prop",
      "bold=true",
      "--prop",
      "spaceAfter=12pt"
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
      "/workspace/review.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Revenue grew 18% year-over-year, ahead of plan.",
      "--prop",
      "size=11pt",
      "--prop",
      "spaceAfter=8pt"
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
      "/workspace/review.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Key Drivers",
      "--prop",
      "style=Heading2",
      "--prop",
      "size=14pt",
      "--prop",
      "bold=true",
      "--prop",
      "spaceBefore=12pt",
      "--prop",
      "spaceAfter=6pt"
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
      "/workspace/review.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Enterprise renewals, upsell, and a new EMEA region.",
      "--prop",
      "size=11pt"
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
      "/workspace/review.docx",
      "/",
      "--type",
      "footer",
      "--prop",
      "type=default",
      "--prop",
      "size=9pt",
      "--prop",
      "text=Page ",
      "--prop",
      "field=page"
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
      "/workspace/review.docx",
      "/footer[1]/p[1]",
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
    "command_name": "save",
    "command_arguments": [
      "/workspace/review.docx"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": [
      "/workspace/review.docx"
    ]
  }
}
```

已验证：`validate` 返回 `no errors found`；`get /footer[1] depth 3` 显示由 5 个 run 组成的 PAGE field chain（begin / instrText / separate / cached value / end）。

## 阅读与分析

先宽后窄。`outline` 展示已有内容；明确位置后，再使用 `view text` / `get` / `query` 深入查看。

Run `view`（`command_arguments` 依次为 path + mode）：`outline`（heading tree / section / TOC / watermark / tracked-changes）→ `html`（batch 编辑后的首轮视觉检查）→ `text start 1 end 80`（内容 QA；path 显示为 `[/body/p[N]]`）→ `annotated` → `stats` → `issues`。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": [
      "/workspace/doc.docx",
      "outline"
    ]
  }
}
```

`watch` 会持续运行 live preview，供人工用户按需打开；agent 自检使用 `view html`。最终的视觉验证应由用户在 Word / WPS / Pages 中打开 `.docx` 完成。

**检查单个 element。** 使用 XPath 风格的语义 path（索引从 1 开始）。path 原样传入 JSON（含 `[N]`）。最后一个 element 应使用 `[last()]`（带括号），`[last]` 会报错。需要 machine-readable output 时加 `json`。常见 path：`/`、`/body/p[1]`、`/body/p[1]/r[1]`、`/body/tbl[1]`（`depth 3`）、`/footer[1]`（`depth 3`，查 fldChar）、`/styles/Heading1`、`/numbering`（`depth 2`）。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": [
      "/workspace/doc.docx",
      "/footer[1]",
      "--depth",
      "3"
    ]
  }
}
```

**跨 document 查询。** 使用 CSS-like selector 做系统性检查，而非手动遍历。operator 包括：`=`、`!=`、`~=`（包含）、`>=`、`<=`、`[attr]`（存在）。完整说明见 Help `query`。示例 selector：`paragraph[style=Heading1]`、`p:contains("quarterly")`、`p:empty`、`image:no-alt`、`paragraph[size>=24pt]`、`field[fieldType!=page]`。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": [
      "/workspace/doc.docx",
      "paragraph[style=Heading1]"
    ]
  }
}
```

`query json` 会将结果包装在 `.data.results[]` 中；用返回结果长度计数（不要用 shell `jq`）。

**大型 document。** 用 `view outline` 按 heading 导航，再用 `query` 跳转；不要把整个 body 全部载入上下文。

## 创建和编辑

可用 verb：`add`（新增 element）、`set`（修改 prop）、`remove`、`move`、`swap`、Batch、`raw-set`（最后手段的 XML；仅 Run + **approval**）。一次构建中，90% 都是 paragraph、run、table、少量 image、一个 TOC 和一个 footer。

### Paragraph、run 与 style

paragraph（`p`）是一个块；run（`r`）是其中 character formatting 一致的一段 span。在 `p` 上设置 paragraph-level prop（style、alignment、spacing、indent）；在 `r` 上设置 font / size / color / bold。

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Executive Summary",
      "--prop",
      "style=Heading1",
      "--prop",
      "size=18pt",
      "--prop",
      "bold=true",
      "--prop",
      "spaceAfter=12pt"
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
      "/workspace/doc.docx",
      "/body/p[1]/r[1]",
      "--prop",
      "color=1F4E79"
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
      "/workspace/doc.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Body follows the heading.",
      "--prop",
      "size=11pt",
      "--prop",
      "style=Normal",
      "--prop",
      "spaceAfter=8pt"
    ]
  }
}
```

垂直间距使用 `spaceBefore` / `spaceAfter`，不要堆叠空 paragraph。左缩进使用 `indent=720`（twips）；首行缩进使用 `firstLineIndent=360`，悬挂缩进使用 `hangingIndent=720`。前导空格会触发 `view issues`。

### 表格

table 位于 `/body/tbl[N]`，由 row `tr[N]` 与 cell `tc[N]` 组成。先按 row/column count 创建，再填充内容。

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body",
      "--type",
      "table",
      "--prop",
      "rows=4",
      "--prop",
      "cols=3",
      "--prop",
      "width=100%"
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
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[1]",
      "--prop",
      "header=true",
      "--prop",
      "c1=Quarter",
      "--prop",
      "c2=Revenue",
      "--prop",
      "c3=Growth"
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
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[1]/tc[1]/p[1]/r[1]",
      "--prop",
      "bold=true"
    ]
  }
}
```

行级 `set` 支持 `height`、`header` 与 `c1 / c2 / … / cN` text shortcut（`cN` 可适用于任意 column count）。cell 格式（bold、fill、color）必须设置在 cell 内的 paragraph / run，**不能**设在 row 上。对单个 cell 的 border，在 `tc` 上设置 cell-level `border.*`（例如 `border.bottom="single;6;000000;0"`），或在内部 paragraph 上设置 paragraph-level `pbdr.*`。

**水平分隔线应使用 paragraph bottom border，不要使用单行 table。** table 作为分隔线会渲染出带最小高度的空白框，尤其会破坏 header/footer。请在 paragraph 上使用 `pbdr.bottom`（`STYLE;SIZE;COLOR`）：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body/p[3]",
      "--prop",
      "pbdr.bottom=single;6;2E75B6"
    ]
  }
}
```

### 列表（项目符号、编号、多级）

单级 bullet/number 应在 paragraph 上设置 `listStyle`（`listStyle` 是 paragraph prop，**不是** run prop，这是常见错误）：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=First item",
      "--prop",
      "listStyle=bullet"
    ]
  }
}
```

多级编号（法律文书风格 1 / 1.1 / 1.1.1）依次 `add` `abstractNum` 与 `num`，再让每个 paragraph 引用 `numId`：

这三步会消费前一步生成的 ID，必须依次 Run，并在每次创建后 readback 实际 ID：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": ["/workspace/doc.docx", "/numbering", "--type", "abstractnum", "--prop", "format=decimal"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": ["/workspace/doc.docx", "/numbering", "--type", "num", "--prop", "abstractNumId=0"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": ["/workspace/doc.docx", "/body", "--type", "paragraph", "--prop", "text=Section one", "--prop", "numId=1", "--prop", "ilvl=0"]
  }
}
```

ID 从 0 开始：全新文档中第一个 `abstractNum` 通常为 0，`num` 通常为 1；示例值必须用 readback 的实际 ID 替换。不存在的 `abstractNumId` 会报错。使用 `query 'paragraph[numId>0]'` 验证。level 与 format 选项见 Help `abstractnum` / `num`。

### Tab stop（签名行、leader row）

tab stop 是 paragraph 的 first-class `tab` child；`pos` 支持 `6in` / `6cm` / twips，`val` 的取值为 `left` / `center` / `right`，`leader` 的取值为 `none` / `dot` / `hyphen` / `underscore`。见 Help `tab`。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body/p[1]",
      "--type",
      "tab",
      "--prop",
      "pos=6in",
      "--prop",
      "val=right",
      "--prop",
      "leader=dot"
    ]
  }
}
```

**Leader 注意事项。** 单独设置 `leader=dot` 不会生成点线；只有文本与 tab stop 之间的 run 中包含真实 `<w:tab/>` 字符时才会显示。做法是先定义 stop，再将 `\t` 写入文本：`text="Chapter 1\t12"`。`\t` 会变为 `<w:tab/>`，点线会延伸至右对齐页码。（字面量 `text="Chapter 1 ......... 12"` 也能交付，但真实 tab stop 对齐更稳定。）

### Field（PAGE / NUMPAGES / DATE / MERGEFIELD / REF）

field 是在 render 时计算的 live value。`fieldType` 指定 field；`name` 给出目标（merge name 或 `ref` bookmark）；`format` / `instr` 添加 switch。

| Field | 用途 | 示例 |
|---|---|---|
| `page` | 当前页码 | 在 footer 上使用 `field=page`，或 inline 使用 `fieldType=page` |
| `numpages` | 总页数 | `field=numpages` / `fieldType=numpages` |
| `date` | 当天日期 | `fieldType=date` + `format=yyyy-MM-dd` |
| `mergefield` | template merge token | `fieldType=mergefield` + `name=CustomerName` |
| `ref` | 对 bookmark 的 cross-reference | `fieldType=ref` + `name=bookmarkName` |

完整的 `fieldType` enum（30+ 个值，包括 `pageref`、`seq`、`styleref`、`docproperty`、`createdate` 等）见 Help `field`。**不存在 `fieldInstr` fieldType**；当 typed shortcut 不足以表达时，使用 `instruction` prop 写入 raw field instruction text。picture switch（`MERGEFIELD Amount \# "#,##0.00"`、`DATE \@ "yyyy年MM月"`）应通过 `instruction=…` 传入；`mergefield` 的 `format` prop 会被忽略并产生 warning，应使用 `instruction`。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body/p[3]",
      "--type",
      "field",
      "--prop",
      "fieldType=mergefield",
      "--prop",
      "name=customer_name"
    ]
  }
}
```

Renders «customer_name» — visible placeholder, replaced in Word at mail-merge time.

**MERGEFIELD template：绝不能渲染字面量 placeholder。** 出现在 body 中的 `{{customer_name}}` 或 `$NAME$` 表示 template 失败，收件人会直接看到它。应插入真实 MERGEFIELD（见上方），或仅将字面 token 放入清晰的说明 paragraph。使用 `query 'field[fieldType=mergefield]'` 确认。

**SEQ / PAGEREF / TOC field value。** Office 工具写入时不保存已 render 的 field value；请按 path 的实际需求重新计算：

- **SEQ 编号**（`Figure 1/2/3`）：`set / prop recalcFields=seq` 会按 body document order 统计 SEQ field，并写入 cached value（`evaluated` 会变为 true；switch/format 见 Help `document`）。相对 heading 的 `\s` 以及 header/footer 内的 SEQ 仍由 Word 处理。
- **PAGE / PAGEREF / NUMPAGES / TOC page number** 依赖分页，而 Office 工具没有 pagination engine；Run `set /settings prop updateFields=true`，让 Word 在打开时计算它们。

多图 document 应同时使用两种方式。学术论文见 `academic-paper` Skill。

### Header 与 footer（页码）

单命令模式：会注入 `<w:fldChar>`，无需手写 field：

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/",
      "--type",
      "footer",
      "--prop",
      "type=first",
      "--prop",
      "text="
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
      "/workspace/doc.docx",
      "/",
      "--type",
      "footer",
      "--prop",
      "type=default",
      "--prop",
      "align=center",
      "--prop",
      "size=9pt",
      "--prop",
      "text=Page ",
      "--prop",
      "field=page"
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
      "/workspace/doc.docx",
      "/footer[2]/p[1]",
      "--prop",
      "align=center"
    ]
  }
}
```

Empty first-page footer — auto-enables differentFirstPage so the cover has no page number. 两者都存在时，default footer 是 `/footer[2]`；只有它时则是 `/footer[1]`。**验证：**`get depth 3` 必须显示 `fldChar` child，而不只是带字面量 `"Page"` 的 run（`view outline` 会为 live field 和 static text 都输出 "Footer: Page"，不可据此判断）。不要执行 `set prop differentFirstPage=true`：该 prop 不受支持，会拒绝，不会静默失败。添加 first-type footer 会自动设置该 bit。复合 **Page X of Y** 见 recipe (b)。

### 目录

任何有 3+ 个 heading 的 document，先确认目标范围内存在真正的 TOC source。TOC source 必须使用 built-in Heading style，或使用带 `outlineLvl` 的 custom paragraph style；仅加粗或放大 Normal text 不是 source。若没有合格 source，应先修正 heading structure，不要插入会在 Word 中显示 `Error! No table of contents entries found` 的 TOC。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/doc.docx",
    "operations": [
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Introduction","style":"Heading1"}},
      {"command":"add","parent":"/styles","type":"style","props":{"id":"ThesisH1","type":"paragraph","outlineLvl":0}},
      {"command":"add","parent":"/body","type":"toc","props":{"levels":"1-3","title":"Table of Contents","hyperlinks":true},"index":0}
    ],
    "stop_on_error": true
  }
}
```

TOC page number 依赖 pagination，Office 工具自身无法计算。必须设置 `updateFields=true`，让 Word 在打开时重新计算 TOC（以及所有 field）；没有 Word-compatible field engine 时，应报告 TOC 为动态且未计算，不得声称 page number 已就绪或猜测静态页码。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/doc.docx",
      "/settings",
      "--prop",
      "updateFields=true"
    ]
  }
}
```

可直接访问 TOC：`/toc[1]` 或 `/tableofcontents` 会解析为第一个 TOC field，供 `get` / `set` / `remove` 使用，无需手动遍历 XPath。

### 图片

picture 位于 run 内。为满足 accessibility 要求，创建时必须直接传入 alt text：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body/p[5]",
      "--type",
      "picture",
      "--prop",
      "src=logo.png",
      "--prop",
      "width=1.5in",
      "--prop",
      "alt=Acme logo"
    ]
  }
}
```

交付前确认 `query 'image:no-alt'` 结果为空。

### 图表

展示数据时，使用**原生 chart**：可编辑、可应用 theme、具备 accessibility，并可在 Word 中重新 render；绝不能用 chart 的平面 PNG screenshot 代替。每个 series 使用 `data="Label:v1,v2,…"`；每个 series 对应一个 `data=`（或使用 `series1=` / `series2=`）。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body",
      "--type",
      "chart",
      "--prop",
      "chartType=bar",
      "--prop",
      "title=Revenue by Region",
      "--prop",
      "categories=EMEA,APAC,Americas",
      "--prop",
      "data=2026:120,150,180"
    ]
  }
}
```

`chartType` 可为 `bar` / `column` / `line` / `pie` / `area` / `scatter`（axis、legend、series style 见 Help `chart`）。只有 Office 工具无法构建的特殊 chart 才可退回为 `type=picture` 的 PNG。

### 超链接和 bookmarks

外部链接通过 `hyperlink`：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body/p[2]",
      "--type",
      "hyperlink",
      "--prop",
      "url=https://example.com",
      "--prop",
      "text=our site"
    ]
  }
}
```

**内部链接**（到 bookmark）使用 `anchor=bookmarkName` — 而不是 `url` 中的 `#fragment`：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body/p[2]",
      "--type",
      "hyperlink",
      "--prop",
      "anchor=chapter1",
      "--prop",
      "text=See Chapter 1"
    ]
  }
}
```

另一种选择是将 `PAGEREF` field 与可见文本配对。请参阅 Help `hyperlink` / `bookmark`。

### 节与页面设置

document root `/` 包含 page setup（`pageWidth`、`pageHeight`、margin，单位为 twips）。多 section document（如横向插页、column）需要 `add` `section` break，见 Help `section`。支持 camelCase（规范写法，如 `pageWidth`）与 lowercase alias（如 `pagewidth`），优先使用 camelCase。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/doc.docx",
    "operations": [
      {"command":"set","path":"/","props":{"pageWidth":12240,"pageHeight":15840,"marginTop":1440,"marginLeft":1440}},
      {"command":"set","path":"/","props":{"columns":2,"columnSpace":720}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Introduction","style":"Heading1","pageBreakBefore":true}}
    ],
    "stop_on_error": true
  }
}
```

Newspaper-style multi-column flow（`columnSpace` in twips；720 = 0.5in）。

### 强制分页

每个逻辑边界只使用一种分页机制。默认是在目标 heading 上设 `pageBreakBefore=true`；也可以在 heading 前插入一个显式 `pagebreak`。不要同时使用两者：会产生双分页，进而出现空白页；在刚添加 `pagebreak` 后的 `p[last()]` 上也不要再设 `pageBreakBefore`。

`break=newPage` 是 `pageBreakBefore=true` 的简写 alias（接受 `newPage|page|nextPage|pageBreak`）。使用 `view html` 预览并检查是否出现空白页。

### 报告级 recipe

以下模式常见于长篇报告，且均已实际执行并通过 `validate`。

**(a) 内容丰富的 cover：达到 ≥ 60% 填充底线。** 依次加入保密 banner、title、subtitle、客户/项目/日期块和 key-theme strip，再将下一 section 强制移至新页：

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/doc.docx",
    "operations": [
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"CONFIDENTIAL — CLIENT USE ONLY","align":"center","size":"9pt","color":"C00000","spaceAfter":"24pt"}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Strategic Growth Review","style":"Title","size":"32pt","bold":true,"align":"center","font":"Cambria","spaceAfter":"8pt"}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"FY26 Outlook and Scenario Planning","italic":true,"size":"16pt","align":"center","spaceAfter":"36pt"}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Prepared for: Acme Corp. Leadership Team","align":"center","size":"11pt"}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Engagement: 2026-04 — 2026-06","align":"center","size":"11pt","spaceAfter":"36pt"}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Key themes: 1) margin resilience, 2) EMEA expansion, 3) capital allocation.","align":"center","italic":true,"size":"10pt"}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Executive Summary","style":"Heading1","pageBreakBefore":true}}
    ],
    "stop_on_error": true
  }
}
```

**(b) Page X of Y footer：组合 PAGE + NUMPAGES。** 先创建 footer paragraph，再通过三个 child operation 构成 live `Page <X> of <Y>`。这是 Help `footer` 的官方 recipe。

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/",
      "--type",
      "footer",
      "--prop",
      "type=default",
      "--prop",
      "text=Page ",
      "--prop",
      "align=center",
      "--prop",
      "size=9pt"
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
      "/workspace/doc.docx",
      "/footer[1]/p[1]",
      "--type",
      "field",
      "--prop",
      "fieldType=page"
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
      "/workspace/doc.docx",
      "/footer[1]/p[1]",
      "--type",
      "run",
      "--prop",
      "text= of "
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
      "/workspace/doc.docx",
      "/footer[1]/p[1]",
      "--type",
      "field",
      "--prop",
      "fieldType=numpages"
    ]
  }
}
```

验证：Run `get /footer[1]/p[1] depth 1`，在返回中统计 `fldChar` 出现次数（expect ≥ 4）。不要用单行 XML 上的“是否匹配”代替计数。

**(c) Header row 使用填充色与白色 bold text。** 顺序很重要：先填入 header cell text（空 cell 中没有 run；对空 cell 执行 `set …/tc[N]/p[1]/r[1]` 会报“未找到”），再设置 cell fill，最后设置 run format：

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body",
      "--type",
      "table",
      "--prop",
      "rows=5",
      "--prop",
      "cols=4",
      "--prop",
      "width=100%"
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
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[1]",
      "--prop",
      "header=true",
      "--prop",
      "c1=Quarter",
      "--prop",
      "c2=Revenue",
      "--prop",
      "c3=Growth",
      "--prop",
      "c4=Status"
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
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[1]/tc[1]",
      "--prop",
      "fill=1F4E79"
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
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[1]/tc[1]/p[1]/r[1]",
      "--prop",
      "bold=true",
      "--prop",
      "color=FFFFFF"
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
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[3]/tc[1]",
      "--prop",
      "fill=D9E2F3"
    ]
  }
}
```

对 header 其余列重复 fill + bold/white；zebra 对 row 3/5 的各列重复 `fill=D9E2F3`（勿用 shell `for`）。

**(d) Financial table：数值右对齐、total bold、total row 的 bottom border。** 对目标 cell 用独立 Run 或一块 Batch 显式 `set`（`align=right`、`bold=true`、`pbdr.bottom="single;6;000000;0"`），勿用循环脚本。

**(e) 含多个 bullet 的 cell（SWOT / risk matrix）。** `c1="a\nb"` 会在**同一个** paragraph 中生成 `<w:br/>` 换行，适合纯多行文本；bullet 则需要独立 paragraph。先用 `set c1=` 写入首项，再在同一 cell 下对每个后续 bullet `add paragraph`（使用 `listStyle=bullet`）：

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[1]",
      "--prop",
      "c1=Installed base of 18k enterprise seats"
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
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[1]/tc[1]",
      "--type",
      "paragraph",
      "--prop",
      "text=Margin structure above peer median",
      "--prop",
      "listStyle=bullet"
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
      "/workspace/doc.docx",
      "/body/tbl[1]/tr[1]/tc[1]/p[1]",
      "--prop",
      "listStyle=bullet"
    ]
  }
}
```

若首个 paragraph 落在底部，使用 `move` 重新排序：`move "/body/tbl[1]/tr[1]/tc[1]/p[N]" index 0`。

**(f) 没有 field engine 时的 TOC。** 纯 pipeline 无法计算可靠的 page number。保留 live TOC、设置 `updateFields=true`，并告知收件人在 Word 或其他兼容 field engine 中打开文档；不要用猜测的静态 page number 替换它。

### Template 交付：分离 Template Notes 与最终用户内容

HR / legal / vendor template 常含仅供内部使用的说明（如“替换 `{{CompanyName}}`”），不得随成品交付。可使用两种方式：

- **末尾的 “Template Notes” section：**放在明确的 `Heading 1`（“Template Notes for HR Users”）下，所有说明置于其后；分发前，从该 heading 起向下 `remove`（先用 `query 'paragraph[style=Heading1]:contains("Template Notes")'` 定位）。
- **由 bookmark 限定的内部 section：**放在 `__template_notes_start` / `_end` bookmark 之间；交付时通过 `raw-set`（Run + **approval**）移除两锚点之间的全部内容。

template 的交付 gate：移除后，`query 'p:contains("Template Notes")'` 和 `query 'p:contains("{{")'` 均必须返回空。若 notes paragraph 仍存在，下游员工会看到内部说明。

### 高级主题（撰写普通报告时可跳过）

报告、memo、letter、proposal 和 HR template 不需要本节。仅当 document 属于 academic（equation、footnote、bibliography）、reviewed（comment、tracked change）或 marked（watermark）时继续阅读。

**Equation 与 footnote。** `type=equation` 接受 LaTeX，支持 `\frac`、`\sum`、Greek letter、`\mathit`、`\mathcal` 等。默认会创建独立的 `/body/oMathPara[N]` display block；若需 inline `<m:oMath>`，请在 paragraph parent path 上使用 `mode=inline`。footnote 按 paragraph index 自动编号。bibliography 的 hanging indent：每个条目使用 `firstLineIndent=-720 indent=720`。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/doc.docx",
    "operations": [
      {"command":"add","parent":"/body","type":"equation","props":{"formula":"\\frac{a}{b} + \\sum_{i=1}^{n} x_i"}},
      {"command":"add","parent":"/body/p[3]","type":"footnote","props":{"text":"See Appendix A for methodology."}},
      {"command":"add","parent":"/","type":"watermark","props":{"text":"DRAFT","color":"BFBFBF","opacity":0.8}}
    ],
    "stop_on_error": true
  }
}
```

**Comment 与 tracked change。** 批量 accept/reject 使用 `set /revision prop revision.action=accept`（或 `reject`）；可用 selector 缩小范围，如 `/revision[@author=Alice]` 或 `/revision[@type=ins]`。使用 `query ins` 与 `query del` 查找单项变更（`trackedchange` 不是 selector）。在 run 上通过 `revision.type=ins|del` + `revision.author=…` 创建 tracked change；完整 `revision.*` set（含 `format` / `moveFrom` / `moveTo`）见 Help `run`。添加 comment：`add "/body/p[4]" type comment`；以 `parentId=N` 建立 reply thread，以 `set "/comments/comment[N]" prop done=true` 标记为 resolved。应 resolve 而非 delete，以保留 audit trail；`query 'comment[done=false]'` 可列出未解决项。prop schema 见 Help `comment` / `run`。

**Watermark。** 使用一条 `add / type watermark`（默认 opacity 为 0.5）；之后可通过 `set /watermark prop opacity=…` 调整。

**何时切换 Skill。** chapter draft、≤ 3 个 footnote、≤ 2 个 equation、没有 bibliography/cross-reference 时继续使用 docx。涉及 citation style（APA / Chicago / IEEE / GB 7714）、in-text ↔ reference 自动链接、含 `\ref` 的编号 equation、“List of Figures”或自动更新 cross-reference 时，切换至 **`academic-paper`**。document 的目标是**数据采集**时，切换至 **`word-form`**，如可填写 form、带用户填写槽的 contract、questionnaire、mail-merge template（`<w:sdt>` content control、`<w:ffData>`、`documentProtection=forms`）。

### raw-set escape hatch（L1 / L2 / L3）

分为三级精度；始终选用能完成任务的最低级别。

- **L1 — high-level prop**（`text=`、`style=Heading1`）：默认方式，覆盖 80% 场景。
- **L2 — dotted-attribute fallback**（`pbdr.top=`、`ind.left=`、`shd.fill=`、`padding.top=`、`font.size=`）：L1 缺少所需设置时使用。例如 `pbdr.bottom="single;6;1F4E79;0"`；会生成 schema-valid XML。
- **L3 — `raw-set` with XML**：最后手段，没有 schema protection。用于 internal hyperlink、composite field 等 typed verb 无法表达的结构（见 XML 附录）。仅逐步 Run + 宿主 **approval**；**永不**进 Batch JSON。

border 格式为 `style;size;color;space`，例如 `single;4;FF0000;1`。hex color 不带 `#`，即 `FF0000`。scheme color name（`accent1..6`、`dark1` / `dark2`、`light1` / `light2`、`hyperlink`）可用于任意接受 hex color 的位置；如需跨 theme 保持稳定，优先使用 hex color。

## QA（必须执行）

**假定存在问题：QA 是找 bug，不是确认步骤。** 第一次生成的 document 几乎不会完全正确；首次检查发现零问题，通常表示检查不够仔细。heading 看似正常，直到 `view outline` 显示 H1 下直接出现 H3；footer 看似显示 “Page 1”，直到 `get depth 3` 发现它其实是 static run 而非 field。

### 声明“完成”前的最小循环

1. Run `view issues`：检查 empty paragraph、missing alt text、formatting anomaly。
2. Run `view outline`：检查 heading hierarchy（不得从 H1 跳到 H3）、TOC 是否存在、section count。
3. Run `view text max-lines 400`：检查 typo、残留的 `\$` / `\t` / `\n` 字面量与 placeholder token。
4. Run `validate`：执行 schema check（Delivery Gate 会在关闭且写入磁盘后的文件上再次执行）。
5. **视觉检查：将整个 document 作为 contact sheet。** 仅适用于具备视觉能力的 agent；若无法解读图像，跳过此步，并在交接时标记 document 为“未做视觉验证”。Run `view screenshot grid auto -o /tmp/sheet.png`，再读取图片。`grid auto` 会将**每一页**平铺在一张图像中（自动选择 column count；`grid 4` 可强制指定），可直观看到 pagination、blank page、heading rhythm、失衡的 margin 与 TOC/cover 位置，而不只是 DOM。Windows+Word 通过真实 Word render 每页，其他平台使用 HTML。若 screenshot 失败，退回 `view html`，并将跨页 break / alignment / rhythm 标记为“未做视觉验证”。thumbnail 只用于**定位**；可疑页应使用无 `grid` 的 `screenshot page N` 以 full resolution 确认细节（column alignment、line spacing、indent、dark-on-dark、caption placement；Windows 使用真实 Word）。`validate` 通过不等于可交付；应达到“看起来像真实 document”的质量。
6. 发现任何问题后先修复，再**重新运行完整循环**；一个修复常会引入另一个问题。

### Delivery Gate（交付前运行；任一失败即 REJECT）

每个 gate 为**独立 Run**（不要 bash / grep / jq）。全部 PASS 前不能声明完成。

**Gate 1 — schema。** Run `close`（若仍 resident），再 Run `validate`。输出须含 `no errors found`；否则 REJECT。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": [
      "/workspace/doc.docx"
    ]
  }
}
```

**Gate 2 — token leak。** Run `view text`，在返回文本中扫描：`$NAME$` 形态、`{{…}}`、`<TODO>`、`xxxx`、`lorem`、字面量 `\$` / `\t` / `\n`。任一命中 → REJECT。TOC placeholder 在 Word-compatible field engine 更新前可接受；改用 structure 确认 TOC field 与 `updateFields`。

**Gate 3 — live PAGE field。** Run `query 'field[fieldType=page]'`（可加 `json`）。结果须 ≥ 1；否则 REJECT（无 live PAGE field）。

全部通过则 Delivery Gate PASS。

### Field / cached-value spot-check

field 的 cached value 在写入时可能过期或为空；应通过**structure 而不是 text**确认其存在。

- **Footer PAGE：**`get /footer[N] depth 3` 会列出 begin / instrText / separate / cached / end 的 run chain：单一 PAGE 至少 5 个 run，组合 “Page X of Y” 至少 11 个。只有一个包含 `"Page"` 的 run 表示 field 缺失；使用 `field=page` 重新添加。
- **TOC：**`get /toc[1] depth 2` 显示 field structure。在重新计算前，page number 可能显示为 `1 1 1 1` 或 `Update field to see…`（见“目录”：设置 `updateFields=true`）。
- **MERGEFIELD：**`query 'field[fieldType=mergefield]'` 应为每个 slot 返回一个结果，其他位置不应存在字面量 `{{name}}`。

### 已知局限

`validate` 只能发现 schema error，不能发现 design error。heading hierarchy 错误、伪 Heading 1 size、placeholder token 被作为 body text、没有 cover 的 document 却带空 first-page footer，都可能通过 `validate`。contact-sheet visual check（`screenshot grid`）与 field structure 检查用于发现这些问题。

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
| `index` 与 `[N]` | `index` 从 0 开始；`[N]` path 从 1 开始 |
| 多次 `add index N` 使用相同 N | 每次插入都会将后续内容下移；复用 N 会让后插入项位于先插入项之前。应逆序插入，或以 `paraId` 为锚点使用 `move after` / `before` |
| path 中的 `[N]` | 原样传入 JSON：`/body/p[1]`（无需 shell 引号） |
| 将 `[last]` 用作 predicate | 必须使用带括号的 `[last()]` |
| spacing 中使用原始 twips | 使用带单位的 value：`12pt`、`0.5cm`、`1.5x` |
| 用空 paragraph 制造 spacing | 使用 `spaceBefore` / `spaceAfter` |
| 对 row 执行 cell format 的 `set` | row `set` 只支持 `height`、`header`、`c1..cN` text；format 应设在 cell paragraph / run |
| 在 run 上设置 `listStyle` | `listStyle` 是 paragraph property |
| 依赖前导空格缩进 | 使用 `indent=720` / `firstLineIndent=360` / `hangingIndent=720`（dotted `ind.left` / `ind.firstLine` 也可） |
| 用 `set differentFirstPage=true` 隐藏 cover page number | 不支持；添加 first-type footer：`type=footer` + `type=first` + `text=""` |
| 需要让新章节另起页 | 在 heading 上使用 `pageBreakBefore=true`；显式 `pagebreak` 只能作为替代方案，绝不能两者同时使用 |
| 把多个 bullet paragraph 合并到一个 cell | `c1="a\nb"` 会得到 `<w:br/>` 换行（同一 paragraph）；独立 bullet paragraph 使用 recipe (e) |
| 何时用 dotted property / `raw-set` | 优先选择 L2 dotted property，L3 `raw-set` 仅作最后手段（Run + approval，不进 Batch） |
| 后一个 paragraph 继承了前一个 heading style | 在下一个 paragraph 显式设置 `style=Normal` |
| 修改仍在 Word 中打开的文件 | 先关闭 Word |
| Batch 中 `$` / 引号 | 使用标准 JSON 字符串；货币 `$` 直接写入 |

## raw-set XML 附录（L3 模式）

`raw-set` 直接注入 literal OOXML，没有 schema protection。仅逐步 Run + 宿主 **approval**；**永不**写入 Batch JSON。`<w:pPr>` 中 element 的顺序为：`pStyle`、`numPr`、`spacing`、`ind`、`jc`、`rPr`（最后）。curly quote 使用 XML entity（`&#x2018;` / `&#x2019;` / `&#x201C;` / `&#x201D;`）。任何带前导或尾随空格的 `<w:t>` 都要添加 `xml:space="preserve"`。RSID 是 8 位 hex number。除非用户另有指定，tracked change 的 author/comment 使用 “Claude”。

**Tracked change 的插入/删除。** 优先在 run 上使用 high-level `revision.type=ins|del`；仅当 typed path 无法表达时才使用 `raw-set`（例如拒绝/恢复其他 author 的变更）。应替换完整的 `<w:r>…</w:r>`，不要把 tag 注入 run 内；复制原 `<w:rPr>` 到两个 run 以保留 format。在 `<w:del>` 内使用 `<w:delText>`（field instruction 则使用 `<w:delInstrText>`）：

```xml
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30</w:delText></w:r></w:del>
<w:ins w:id="2" w:author="Claude" w:date="2026-01-01T00:00:00Z"><w:r><w:t>60</w:t></w:r></w:ins>
<w:r><w:t> days.</w:t></w:r>
```

删除 paragraph / list item 的所有内容时，还要在 paragraph mark（`<w:pPr><w:rPr>`）内写入 `<w:del/>`；否则 accept change 后会留下空 paragraph。要**拒绝另一 author 的插入**，将自己的 `<w:del>` 嵌套在对方的 `<w:ins>` 中；要**恢复对方的删除**，在后方添加 `<w:ins>`，不要修改对方的原始结构。

**Internal hyperlink 到 bookmark**（优先使用 high-level `anchor=` path；`raw-set` 仅用于命令无法表达的自定义 run style）：

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

**Comment marker** 与 `<w:r>` 是 sibling，不能位于其中。reply thread 与 resolved status 属于 high-level 功能，可使用 `parentId=` / `done=`（见上方 comment 说明）：

```xml
<w:commentRangeStart w:id="0"/><w:r><w:t>annotated text</w:t></w:r><w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
```

使用 `set /settings prop updateFields=true` 强制在打开时重新计算 field（写入 `<w:updateFields w:val="true"/>`，适用于依赖 layout 的 PAGE / PAGEREF / NUMPAGES / TOC page number，不需要 `raw-set`）。对 SEQ 编号，优先使用 `set / prop recalcFields=seq`；它会直接写入正确的 cached value，无需等待 Word。

### Help 速查

遇到疑问时使用 `{{OFFICE_HELP_TOOL}}`（`command_arguments`: `docx` / element / verb+element）。Help 是权威 schema，本 Skill 用于帮助决策。
