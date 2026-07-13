---
name: officecli
description: 使用 officecli CLI 创建、分析、校对和修改 Office 文档（.docx、.xlsx、.pptx）。当用户要创建、检查、核验格式、排查问题、添加图表或修改 Office 文档时使用。
---

# officecli

AI 友好的 .docx、.xlsx、.pptx CLI。单一二进制、无依赖，无需安装 Office。

## 策略

**L1（读取）→ L2（DOM 编辑）→ L3（原始 XML）**。始终优先使用较高层级；使用 `--json` 获取结构化输出。

**开始文档工作前，先检查专项 Skill**（见本文末尾）。融资演示文稿、学术论文、财务模型、仪表盘和 Morph 动画都需要先加载对应 Skill：执行一次 `load_skill`，再继续操作。

---

## Help 系统（重要）

**不确定属性名、值格式或命令语法时，务必运行 help，不要猜测。** 一次 help 查询胜过猜测-失败-重试的循环。

`officecli help` 等同于 `officecli --help`；`officecli <cmd> --help` 等同于 `officecli help <cmd>`，两者内容相同。

```bash
officecli help                                  # 全部命令、全局选项和 schema 入口
officecli help docx                             # 列出全部 docx 元素
officecli help docx paragraph                   # 完整 schema：属性、别名、示例、读取结果
officecli help docx set paragraph               # 按动词过滤：只显示可用于 `set` 的 props
officecli help docx paragraph --json            # 结构化 schema（机器可读）
```

格式别名：`word`→`docx`、`excel`→`xlsx`、`ppt`/`powerpoint`→`pptx`。动词：`add`、`set`、`get`、`query`、`remove`。MCP 通过唯一的字符串参数 `command` 暴露同一套 schema：`{"command":"help docx paragraph"}`。不要传入结构化的 `{"format":...,"type":...}` 对象：MCP 工具只有一个 `command` 参数，并会将其原样传给 CLI。

---

## 性能：resident 模式

**每个命令首次访问时都会自动启动 resident**（空闲超时 60 秒），从而自动避免文件锁冲突。较长会话仍建议显式使用 `open`/`close`（空闲超时 12 分钟）：
```bash
officecli open report.docx       # 显式保留在内存中
officecli set report.docx ...    # 没有文件 I/O 开销
officecli close report.docx      # 保存并释放
```

如需禁用自动启动，设置：`OFFICECLI_NO_AUTO_RESIDENT=1`。

**仅在与非 officecli 程序交接的边界执行 flush。** officecli 自身的读取（`get`/`query`/`view`/`dump`）始终能看到最新修改，因此工作流中无需中途保存。仅在**非 officecli 程序读取文件前**运行 `save`（保留 resident）或 `close`（flush 并释放），例如 python-docx/openpyxl、Word、渲染器、交付或上传。（空闲会话会在数秒内自动 flush；`OFFICECLI_RESIDENT_FLUSH=each` 会让每次修改在返回前 flush。）

---

## 快速开始

**PPT：**
```bash
officecli create slides.pptx
officecli add slides.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add slides.pptx '/slide[1]' --type shape --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm --prop font=Arial --prop size=24 --prop color=FFFFFF
```

**Word：**
```bash
officecli create report.docx
officecli add report.docx /body --type paragraph --prop text="Executive Summary" --prop style=Heading1
officecli add report.docx /body --type paragraph --prop text="Revenue increased by 25% year-over-year."
```

**Excel：**
```bash
officecli create data.xlsx
officecli set data.xlsx /Sheet1/A1 --prop value="Name" --prop bold=true
officecli set data.xlsx /Sheet1/A2 --prop value="Alice"
```

---

## L1：创建、读取与检查

```bash
officecli create <file>               # 创建空白 .docx/.xlsx/.pptx（类型由扩展名决定）
officecli view <file> <mode>          # outline | stats | issues | text | annotated | html
officecli get <file> <path> --depth N # 获取节点及其子节点 [--json]
officecli query <file> <selector>     # CSS 风格查询
officecli validate <file>             # 根据 OpenXML schema 验证
```

### view 模式

| 模式 | 说明 | 常用参数 |
|------|-------------|-------------|
| `outline` | 文档结构 | |
| `stats` | 统计信息（页数、字数、shape 数量） | |
| `issues` | 格式、内容或结构问题 | `--type format\|content\|structure`、`--limit N` |
| `text` | 纯文本提取 | `--start N --end N`、`--max-lines N` |
| `annotated` | 带格式标注的文本 | |
| `html` | 静态 HTML 快照：与 `watch` 使用同一渲染器，无需 server | `--browser`、`--page N`（docx）、`--start N --end N`（pptx） |
| `screenshot` / `svg` / `forms` | 通过 headless browser 生成 PNG / SVG（pptx 幻灯片）/ 表单字段 JSON（docx） | `-o`、`--screenshot-width/-height`、pptx `--grid N` |

一次性快照（CI artifact、归档、diff）使用 `view html`；需要实时刷新或在 browser 中点击选择时使用 `watch`。

### get

通过元素 localName 可访问任意 XML 路径。使用 `--depth N` 展开子节点；添加 `--json` 获取结构化输出。默认文本输出适合 grep：`path (type) "text" key=val key=val ...`。

```bash
officecli get report.docx '/body/p[3]' --depth 2 --json
officecli get slides.pptx '/slide[1]' --depth 1          # 列出第 1 张幻灯片的全部 shape
officecli get data.xlsx '/Sheet1/B2' --json
```

### 稳定 ID 寻址

具有稳定 ID 的元素会返回 `@attr=value` 路径，而非位置索引。多步骤工作流应优先使用它们：位置索引会随插入/删除变化，稳定 ID 不会。

```
/slide[1]/shape[@id=550950021]                    # PPT shape
/slide[1]/table[@id=1388430425]/tr[1]/tc[2]       # PPT table
/body/p[@paraId=1A2B3C4D]                         # Word paragraph
/comments/comment[@commentId=1]                    # Word comment
```

PPT 也接受 `@name=`（例如 `shape[@name=Title 1]`），并能识别 morph 的 `!!` 前缀。没有稳定 ID 的元素（slide、run、tr/tc、row）会回退到位置索引。

### query

CSS 风格 selector：`[attr=value]`、`[attr!=value]`、`[attr~=text]`、`[attr>=value]`、`[attr<=value]`、`:contains("text")`、`:empty`、`:has(formula)`、`:no-alt`。`query`/`set`/`remove` 支持布尔 `and`/`or`：`cell[value>5000 or value<100]`、`cell[(type=Number or type=Date) and value>0]`。Excel 可按列名筛选行：`Sheet1!row[Salary>5000]`。`set` 同时接受 selector 和 Excel 原生路径（与 `get`/`query` 保持一致）。`set`/`remove` 不接受未限定范围的裸 selector。

```bash
officecli query report.docx 'paragraph[style=Normal] > run[font!=Arial]'
officecli query slides.pptx 'shape[fill=FF0000]'
```

---

## Watch 与交互式选择

实时 HTML 预览：每次文件变化都会自动刷新。可在 browser 中点击、按住 shift 点击或框选 shape；CLI 能读取当前 browser 选择并据此操作。

```bash
officecli watch <file> [--port N]      # 启动预览 server（默认端口 26315）
officecli unwatch <file>               # 停止
officecli goto <file> <path>           # 将正在观察的 browser 滚动至元素（docx：p / table / tr / tc）
```

打开输出的 `http://localhost:N` URL。单击选择；shift/cmd/ctrl+单击多选；从空白处拖动可框选。PPT/Word 使用蓝色轮廓；Excel 使用原生风格绿色选择（双击 cell 可行内编辑；拖动 chart 可调整位置）。

### `get <file> selected`：读取用户点击的内容

```bash
officecli get <file> selected [--json]
```

返回当前选择项的 DocumentNodes。未选择任何内容时返回空结果；未运行 watch 时退出码不为 0。

```bash
# 用户在 browser 中点击 shape 后，要求“把这些变红”
PATHS=$(officecli get deck.pptx selected --json | jq -r '.data.Results[].path')
for p in $PATHS; do officecli set deck.pptx "$p" --prop fill=FF0000; done
```

### 关键属性

- **选择在文件编辑后仍会保留。** 路径使用稳定的 `@id=` 形式。
- **所有已连接 browser 共享一个选择集。** 后写入者生效。
- **同一文件只能有一个 watch。** 每个文件同一时间只能运行一个 watch process。
- **group shape 整体选择。** v1 不支持深入选择 group 内的单个子元素。
- **覆盖范围：**`.pptx` 的 shape/picture/table/chart/connector/group；`.docx` 的顶层 paragraph 和 table。继承的 layout/master 装饰，以及 Word 嵌套元素（table cell、run 级）不可寻址。**`.xlsx` 不会输出 `data-path`**，因此 xlsx 上的 `mark`/`selection` 始终解析为 `stale=true`（v2 候选能力）。

### Mark：等待审核的编辑提案

需要人工审核后才能写入文件的修改，请使用 `mark`。Mark 仅存在于 watch process 中；通过独立的 `set` pipeline 应用已接受的修改。一次性修改直接使用 `set`；需要永久写入文件的批注使用 `add --type comment`（Word 原生）。

```bash
officecli mark <file> <path> [--prop find=... color=... note=... tofix=... regex=true] [--json]
officecli unmark <file> [--path <p> | --all] [--json]
officecli get-marks <file> [--json]
```

Props：`find`（字面文本，或在 `regex=true` 时使用 regex；原始形式为 `find='r"[abc]"'`）、`color`（hex / `rgb(...)` / 22 个命名白名单颜色）、`note`、`tofix`（驱动应用 pipeline）。**Path** 必须使用 watch HTML 中的 `data-path` 格式；完整 pipeline 见子 Skill。

---

## L2：DOM 操作

### set：修改属性

```bash
officecli set <file> <path> --prop key=value [--prop ...]
```

通过元素路径（由 `get --depth N` 找到）可设置**任何 XML 属性**，包括当前不存在的属性。未提供 `find=` 时，`set` 会将格式应用于整个元素。

**值格式：**

| 类型 | 格式 | 示例 |
|------|--------|---------|
| 颜色 | Hex（可带或不带 `#`）、命名颜色、RGB、theme | `FF0000`、`#FF0000`、`red`、`rgb(255,0,0)`、`accent1`..`accent6` |
| 间距 | 带单位 | `12pt`、`0.5cm`、`1.5x`、`150%` |
| 尺寸 | EMU 或带后缀 | `914400`、`2.54cm`、`1in`、`72pt`、`96px` |

**点号属性别名**：shape/run/paragraph/table/row/cell/section/style 接受 `font.<attr>` 形式，例如 `--prop font.color=red --prop font.bold=true --prop font.size=14pt`。完整列表请运行 `officecli help <fmt> <element>`。

### find：格式化或替换匹配文本

在 `set` 上使用顶层 `--find` / `--replace`（`query` 上使用 `--find`）。旧式 `--prop find=X` 仍可用，但会输出提示。

```bash
# 格式化匹配文本（自动拆分 run）
officecli set doc.docx '/body/p[1]' --find weather --prop bold=true --prop color=red

# Regex 匹配（regex= 仍是 prop 标记）
officecli set doc.docx '/body/p[1]' --find '\d+%' --prop regex=true --prop color=red

# 替换文本（使用 `/` 表示整份文档范围）
officecli set doc.docx / --find draft --replace final

# docx：带修订跟踪的查找替换
officecli set doc.docx / --find draft --replace final --prop revision.author=Alice

# PPT：语法相同，路径不同
officecli set slides.pptx / --find draft --replace final
```

**Path 控制搜索范围：**`/` = 整份文档，`/body/p[1]` 或 `/slide[N]/shape[M]` = 指定元素，`/header[1]` / `/footer[1]` = 页眉/页脚。

**注意：**
- 默认区分大小写。不区分大小写：`--prop 'find=(?i)error' --prop regex=true`
- 匹配可跨 run 边界
- 未匹配时静默成功。`--json` 包含 `"matched": N`
- **Excel：**仅支持 `find` + `replace`（不支持 find + 格式 props）

### add：添加元素或克隆

```bash
officecli add <file> <parent> --type <type> [--prop ...]
officecli add <file> <parent> --type <type> --after <path> [--prop ...]   # 在锚点后插入
officecli add <file> <parent> --type <type> --before <path> [--prop ...]  # 在锚点前插入
officecli add <file> <parent> --type <type> --index N [--prop ...]        # 从 0 开始的位置（旧式）
officecli add <file> <parent> --from <path>                               # 克隆现有元素
```

`--after`、`--before`、`--index` 互斥。未提供位置选项时追加到末尾。

**元素类型（含别名）：**

| 格式 | 类型 |
|--------|-------|
| **pptx** | slide（含 hidden）、shape（font.latin/ea/cs、direction=rtl、underline.color、highlight=COLOR，支持 Add/Set/Get/HTML preview，effective.X+effective.X.src；`arrow` 是 `rightArrow` 的别名；slideMaster/slideLayout 支持按类型 add/set/remove）、picture（SVG、brightness/contrast/glow/shadow、rotation、link、tooltip）、chart（direction=rtl、pieOfPie、barOfPie、axisLine/gridline 逐属性 setter、animation+chartBuild=byCategory\|bySeries、line 的 dropLines/hiLowLines/upDownBars、`anchor=x,y,w,h` 简写）、table（cell direction=rtl、fill/background、内置 PowerPoint style 目录、`/col[C]` get + swap/copyFrom、row/col Move/CopyFrom）、row（tr）、connector（from/to 接受完整路径 `@name=`/`@id=` 形式；拒绝裸 `@name=Foo`，必须使用 `/slide[N]/shape[@name=Foo]`；startshape/endshape 支持 SetByPath；默认边到边锚定，fromSide/toSide 强制指定边，fromIdx/toIdx 指定原始 cxn 索引）、group（link、tooltip，可通过 get/query/add/remove 深度遍历，`ungroup=true` 会还原为 slide-absolute）、align/distribute（targets= 接受 shape[@id=N] 路径而非仅位置索引）、video/audio（loop、autoStart 别名）、equation、notes（direction=rtl、lang）、comment（legacy + modern p188 threaded round-trip）、animation（15 个 emphasis + 16 个 exit preset、多效果链、motion-path preset、repeat/restart/autoReverse、chart animation）、transition（12 个 p15 preset + morph/p14）、paragraph（para）、run、zoom、ole（preview=，通过 add-part+raw-set 支持完整 dump round-trip）、placeholder（phType=...）、model3d（rotation=ax,ay,az；支持完整 dump round-trip）、smartart（通过 add-part 支持 dump round-trip）、diagram（仅 add：mermaid → 原生 shape 或渲染图片，`--type diagram`/`flowchart`）。 |
| **docx** | paragraph（direction/font.latin/ea/cs、bold.cs/italic.cs/size.cs、lang.latin/ea/cs、wordWrap、framePr.\*、tabs 简写）、run（lang slot、direction、underline.color、position half-pts，**revision.type=ins\|del\|format\|moveFrom\|moveTo + revision.action=accept\|reject**，并可使用 .author/.date；在 `set /revision[...]` 上可使用裸 `@author=`/`@type=` selector 筛选接受/拒绝，但 `query 'revision[...]'` 必须使用点号形式 `revision.author=`/`revision.type=`；move+revision 仅适用于 run 级路径，不适用于 paragraph 级；paragraph/shape 路径的 **range=START:END** 使用显式 0 起始、半开区间的字符偏移格式化字符片段，而非寻址 run，是与 find= 对等的偏移查找方式）、table（direction=rtl、hMerge，row 上的 cantSplit/cell 上的 nowrap，均支持 add+set；**虚拟列操作**：在 /body/tbl[N]/col 上 add/remove/move/copyfrom）、row（tr）、cell（td）、image、header/footer（direction）、section（pageNumFmt 完整 enum、direction=rtl、rtlGutter、pgBorders=box）、bookmark、comment、footnote、endnote、formfield、sdt、chart、equation、field（28 种类型）、hyperlink、style（Add/Set 中的 direction、indents、pbdr、lineSpacing）、toc、watermark、break、ole、**num/abstractNum/lvl**、**tab**、**textbox/shape**（主要支持 add：Get 仅返回原始 XML preview，不提供结构化 readback；Set 仅限 width/height/geometry/fill/line.\*；位置使用 `anchor.x`/`anchor.y`，不是裸 x/y；仅 **textbox** 支持 `textDirection`/rotation/gradient/shadow，docx shape 本身不支持 rotation 或 gradient）、嵌入式 **OLE 可通过 dump→batch round-trip**、**diagram**（仅 add：mermaid → 原生 shape 或渲染图片，`--type diagram`/`flowchart`；add 时不支持 x/y，通过 `set /body/group[N]` 重新定位）。支持 docDefaults.rtl、autoHyphenation；`get /` 输出 locale + /comments /footnotes /endnotes。`create --minimal` 用于原始 OOXML scaffold。 |
| **xlsx** | sheet（visible/hidden/veryHidden、print margins、printTitleRows/Cols、rightToLeft sheetView、可感知级联的 rename）、row（`c{N}=` 为 cell-content 简写；add 接受 `--from /Sheet/col[L]`；插入时重写 formula reference）、col（重写 formula reference，移动时 named-range 跟随）、cell（type=richtext+runs、merge=range/sweep、direction=rtl、phonetic；**remove 时 `--shift left\|up`，add 时 `shift=right\|down`**，与 Excel UI dialog 对齐；自动识别 formula；calc 中支持 OFFSET/INDIRECT）、chart（按 axis 设置 RTL/title、anchor=x,y,w,h、pareto）、image（SVG）、comment（direction=rtl）、table（listobject）、namedrange（definedname、volatile、`[@name=X]`；解析时会内联 formula body）、pivottable（cache CoW + 跨 pivot 共享，labelFilter=field:type:value 仅 add 时可用，topN=integer 仅 add 时可用，fillDownLabels 是 repeatLabels 的别名而非独立能力，calculatedField）、sparkline、validation、autofilter、shape、textbox、CF（databar/colorscale/iconset/formulacf/cellIs/topN/aboveAverage）、ole、csv。Query 支持 `merge`/`mergedrange`。Workbook 支持 password。Shape selector 会枚举 grpSp 内的叶子节点。 |

### Pivot table（xlsx）

```bash
officecli add data.xlsx /Sheet1 --type pivottable \
  --prop source="Sheet1!A1:E100" --prop rows=Region,Category \
  --prop cols=Year --prop values="Sales:sum,Qty:count" \
  --prop grandTotals=rows --prop subtotals=off --prop sort=asc
```

关键 props：`rows`、`cols`、`values`（Field:func[:showDataAs]）、`filters`、`source`、`position`、`layout`（compact/outline/tabular）、`repeatLabels`、`blankRows`、`aggregate`、`showDataAs`（percent_of_total/row/col、running_total）、`grandTotals`、`subtotals`、`sort`。聚合器：sum、count、average、max、min、product、stdDev、stdDevp、var、varp、countNums。日期列会自动分组。完整 schema 请运行 `officecli help xlsx pivottable`。

### 文档级属性（全部格式）

```bash
officecli set doc.docx / --prop docDefaults.font=Arial --prop docDefaults.fontSize=11pt
officecli set doc.docx / --prop protection=forms --prop evenAndOddHeaders=true
officecli set data.xlsx / --prop calc.mode=manual --prop calc.refMode=r1c1
officecli set slides.pptx / --prop defaultFont=Arial --prop show.loop=true --prop print.what=handouts
```

全部文档级属性（docDefaults、docGrid、CJK spacing、calc、print、show、theme、extended）请运行 `officecli help <format> /`。

### 排序（xlsx）

```bash
officecli set data.xlsx /Sheet1 --prop sort="C desc" --prop sortHeader=true
officecli set data.xlsx '/Sheet1/A1:D100' --prop sort="A asc" --prop sortHeader=true
```

格式：`COL DIR[, COL DIR ...]`。包含 merged cell 或 formula 的范围会被拒绝。Sidecar metadata（hyperlink、comment、conditional formatting、drawing）会自动随行移动。

### 文本锚定插入（`--after find:X` / `--before find:X`）

通过 paragraph 内的文本匹配定位插入点。inline 类型（run、picture、hyperlink）会插入 paragraph 内；block 类型（table、paragraph）会自动拆分 paragraph。PPT 仅支持 inline 类型。

```bash
# Word：在匹配文本后插入 inline run
officecli add doc.docx '/body/p[1]' --type run --after find:weather --prop text=" (sunny)"

# Word：在匹配文本后插入 block table（自动拆分 paragraph）
officecli add doc.docx '/body/p[1]' --type table --after "find:First sentence." --prop rows=2 --prop cols=2
```

### 克隆

`officecli add <file> / --from '/slide[1]'`：会复制全部跨 part relationship。

### move、swap、remove

```bash
officecli move <file> <path> [--to <parent>] [--index N] [--after <path>] [--before <path>]
officecli swap <file> <path1> <path2>
officecli remove <file> '/body/p[4]'
```

使用 `--after` 或 `--before` 时可省略 `--to`，目标容器会从锚点推断。

### batch：在一个保存周期内执行多项操作

默认在发生错误后继续执行（任一 item 失败时返回 exit 1）。使用 `--stop-on-error` 在第一次失败时终止。`--force` 用于绕过 docx protection。

`officecli dump <file> [<path>]` 会输出可重放的 batch JSON，用于 round-trip：`.docx`（完整覆盖）、`.pptx`（text/table/picture/chart/note/theme，以及通过 raw-set passthrough 的 OLE/3D/video/audio/SmartArt/morph/p15 transition）、`.xlsx`（cell/formula/style 以及 table、conditional formatting、validation、comment、chart、sparkline、picture、shape、pivot table；slicer/chartEx/OLE 通过 verbatim carrier）。Path 默认是 `/`（整份文档）；可传入子树路径缩小 dump 范围（docx：`/body`、`/body/p[N]`、`/body/tbl[N]`、`/theme`、`/settings`、`/numbering`、`/styles`；xlsx：`/SheetName`、`/sheet[N]`）。重放后，`officecli refresh <file.docx>` 会重新计算 TOC 页码 / PAGE / cross-reference（Windows 使用 Word backend，其他环境使用 headless-HTML fallback）。

```bash
echo '[
  {"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
  {"command":"set","path":"/Sheet1/B1","props":{"value":"Score","bold":"true"}}
]' | officecli batch data.xlsx --json

officecli batch data.xlsx --commands '[{"op":"set","path":"/Sheet1/A1","props":{"value":"Done"}}]' --json
officecli batch data.xlsx --input updates.json --force --json
```

支持：`add`、`set`、`get`、`query`、`remove`、`move`、`swap`、`view`、`raw`、`raw-set`、`validate`。字段：`command`（或 `op`）、`path`、`parent`、`type`、`from`、`to`、`index`、`after`、`before`、`props`、`selector`、`mode`、`depth`、`part`、`xpath`、`action`、`xml`。

---

## L3：原始 XML

当 L2 无法表达所需操作时使用。无需声明 xmlns，prefix 会自动注册。

```bash
officecli raw <file> <part>                          # 查看原始 XML
officecli raw-set <file> <part> --xpath "..." --action replace --xml '<w:p>...</w:p>'
officecli add-part <file> <parent>                   # 创建新的 document part（返回 rId）
```

`raw-set` action：`append`、`prepend`、`insertbefore`、`insertafter`、`replace`、`remove`、`setattr`。可用 part 请运行 `officecli help <format> raw`。

---

## 常见陷阱

| 常见错误 | 正确做法 |
|---------|-----------------|
| `--name "foo"` | 使用 `--prop name="foo"`：全部属性都通过 `--prop` 传入 |
| 在 zsh/bash 中使用未加引号的 `[N]` 路径 | 始终加引号：`'/slide[1]'` 或 `"/slide[1]"`（shell 会对方括号执行 glob 展开） |
| 用 PPT `shape[1]` 承载内容 | `shape[1]` 通常是标题 placeholder；内容 shape 请使用 `shape[2]+` |
| `/shape[myname]` | 不支持按名称索引。使用数字索引或 `@name=`（仅 PPT） |
| 猜测属性名称 | 运行 `officecli help <format> <element>` 查看准确名称 |
| 修改已打开的文件 | 先在 PowerPoint/WPS 中关闭该文件 |
| shell 字符串中的 `\n` | 在 `--prop text="..."` 中使用 `\\n` 表示换行 |
| shell 文本中的 `$` | `--prop text="$15M"` 会移除 `$15`。使用单引号：`--prop text='$15M'`，或使用 heredoc batch |

---

## 专项 Skill

`officecli load_skill <name>` 会输出一个 SKILL.md，请遵循其中规则。

**加载规则：**
- 在“适用场景”中选择最具体的匹配项；没有匹配时，加载格式默认 Skill（`word` / `pptx` / `excel`）。
- 场景 Skill 已包含格式默认规则：每个 artifact **只加载一个** Skill，不要叠加。
- 已加载的规则会跨 turn 保留，无需每次回复重新加载。
- 两个不同 artifact → 分别加载两个 Skill。

### Word（.docx）

| 名称 | 适用场景 |
|------|-------------|
| `word` | 报告、信函、备忘录、提案和通用文档 |
| `academic-paper` | 期刊、会议论文、论文：APA / Chicago / IEEE / MLA 引文、公式、SEQ + PAGEREF cross-reference、多栏期刊布局、参考文献。**不用于**商务报告或信函（应使用 `word`） |

### PowerPoint（.pptx）

| 名称 | 适用场景 |
|------|-------------|
| `pptx` | 通用 deck：board review、sales deck、all-hands、product launch |
| `pitch-deck` | **仅融资**：seed / Series A-C / SAFE / convertible / strategic raise。**不用于** sales / product / board deck（应使用 `pptx`） |
| `morph-ppt` | 电影感的 Morph 动画演示文稿。**不用于**静态 deck（应使用 `pptx`） |
| `morph-ppt-3d` | 3D Morph：GLB model、camera move、depth。**不用于**仅 2D 的 Morph（应使用 `morph-ppt`） |

### Excel（.xlsx）

| 名称 | 适用场景 |
|------|-------------|
| `excel` | 通用 workbook、formula、pivot、tracker |
| `financial-model` | 财务模型、情景分析、预测。**不用于**一般数据分析（应使用 `excel`） |
| `data-dashboard` | CSV/表格数据 → 带 chart 和 sparkline 的 KPI / analytics / executive dashboard。**不用于**原始数据追踪（应使用 `excel`） |

示例：融资 deck 任务 → `officecli load_skill pitch-deck` → 使用输出的规则。

---

## 注意事项

- 路径为**从 1 开始**（XPath 约定）：`'/body/p[3]'` = 第 3 个 paragraph
- `--index` 为**从 0 开始**（array 约定）：`--index 0` = 第一个位置
- **Excel 例外：**`add --type row` 和 `add --type col` 的 `--index N` **从 1 开始**（与 OOXML RowIndex / 列字母索引一致）。`--index 5` 插入到第 5 行 / 第 5 列。
- 修改后使用 `validate` 和/或 `view issues` 验证
- **不确定时**，运行 `officecli help <format> <element>`，不要猜测
