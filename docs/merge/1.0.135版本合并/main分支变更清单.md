# 1.0.135 版本合并：main 分支变更清单

## 1. 对比口径

- 共同基线：`fe58b0b0`。
- 合并前的 `dev`：`f7a21528`。
- 合并时的 `main`：`25382af4`。
- 当前合并提交：`30437b72`。
- 本文的“已合并”表示代码已经进入当前 `dev` 工作区（其中本轮新增内容尚待提交）；“部分合并”表示底层代码存在，但 CLI 入口或完整链路没有接上；“未合并”表示 main 的功能代码没有进入当前 dev。
  已提交的合并基线：`30437b72`
- `main` 从共同基线之后共有约 191 个提交，其中 188 个是非 merge 提交。多数提交是 Excel 兼容性修复，不是完全独立的新功能。
  Git 提交：`25382af4`, `fe58b0b0`

## 2. 总体结论

| 方向 | 大白话总结 | 当前状态 | 关键提交 |
| --- | --- | --- | --- |
| Excel 数据可靠性 | 错误输入尽早拒绝，失败写入回滚，公式、日期、表格、图表关联更可靠 | 已合并 | `e9197b87`, `dd065553`, `2da48a90`, `2b1aa731` |
| 查询和 selector | 查询条件、错误提示、批量命令和未知参数处理更准确 | 已合并 | `33c851e0`, `06abe02f`, `3a5e4772` |
| Word/PPTX/渲染 | Word 预览和 round-trip 更稳定，PPTX 文本适配修复 | 已合并 | `bfccf47d`, `69a04774`, `4e1ed33f` |
| Resident 和核心稳定性 | 单文件只允许一个 resident，自动保存和子进程通信更稳 | 已合并 | `2285bfcf`, `8062abc8`, `f7acea3a` |
| `query --compact --fields` | 面向 AI/脚本的一行一元素输出，可选择追加格式字段 | 已合并（待提交） | `ae7fcd54`, `16c7c6ae` |
| `view --range` | 只查看 Excel 小区域，或只截 PPTX/Word 某个元素 | 已合并（待提交） | `f6272850`, `d9cfa304`, `4cde9aad` |
| 插件生态和 PDF | 支持额外文件格式、插件预览和 PDF 导出 | 未合并 | `2b32d4b3`, `b439d14e`, `fa2bfd9f` |
| 安装、更新、skills install | 自动安装 binary、skill、MCP，后台检查和升级 | 未合并 | `1ed7b355`, `ab52656d`, `6ca4486d` |
| SDK 自动安装 | Node/Python SDK 找不到 CLI 时自动下载或安装 | 未合并 | `1648b007`, `6ce96cd3`, `186a0e5e` |
| 测试和 CI | 已恢复 main 历史测试，并保留 dev 手工测试；安装 Smoke Test 和平台兼容性检查未合并 | 部分合并 | `020fc7a1`, `c08ebfc7`, `65893f46` |

## 3. 已合并：对产品有直接价值的变更

### 3.1 Excel 输入校验和失败回滚

1. **错误值不再悄悄写入。** 数字、公式长度、公式参数数量、XML 非法字符、布尔值和图表轴参数不符合要求时，直接返回错误，不再生成一个“看似成功、实际有问题”的文件。
   Git 提交：`e9197b87`, `b0b44ad5`, `7780a585`, `eb583d93`, `775f23b1`

2. **写到一半失败会撤销半成品。** Add cell、Add chart、Add shape、Add slicer、Set chart 等操作在后续校验失败时，会尽量恢复原状态，避免留下孤立 XML 部件或幽灵 cell。
   Git 提交：`dd065553`, `563e10ff`, `cefe37c2`, `81f4a28d`, `846cfb9b`

3. **非法的范围和坐标会明确报错。** 行列范围、绘图 anchor、条件格式范围、数据验证范围、表格范围不再被静默截断或错误转换。
   Git 提交：`0a0bb7bb`, `b88a04c6`, `0f14fe82`, `1852f82e`, `f9bac521`

### 3.2 Excel 日期、数字和公式

1. **支持 1904 日期系统。** 老 Mac Excel 文件的日期不会再整体偏移。
   Git 提交：`2da48a90`, `22d16848`

2. **数字、布尔值和格式读写更准确。** 布尔值会读成 `TRUE/FALSE`，日期、百分比、货币和自定义格式不再轻易被当成普通字符串或错误数字。
   Git 提交：`7780a585`, `27d6c157`, `f4317f35`, `c23debd0`

3. **公式缓存不再长期过期。** 保存时会刷新可计算公式；公式太复杂或扫描被打断时，下一次空闲会继续处理，不会永久保存半更新状态。
   Git 提交：`dcdb6f64`, `d3a52e8d`, `1663a241`, `46d42ad9`

4. **公式引用移动更准确。** 插入、删除行列或重命名 sheet 后，跨 sheet 公式、`#REF!` 状态和缓存会同步处理。
   Git 提交：`bb5b26fc`, `25837261`, `46d42ad9`

### 3.3 Excel 表格、行列和结构关联

1. **插入/删除行列后，表格关联内容跟着走。** 表头、tableColumns、自动筛选、条件格式、冻结窗格、图表引用都会同步更新。
   Git 提交：`2b1aa731`, `df90233d`, `680c84b9`, `d1a32ddb`, `808cf5eb`

2. **表格总计行和表头状态更可靠。** 开关 totalRow、headerRow 时，会清理旧的总计单元格和自动筛选状态，避免 Excel 打开时修复文件。
   Git 提交：`3497d1d5`, `24e255c8`, `fb392e00`, `8cca83b2`

3. **复杂表头更容易操作。** 表头包含空格、逗号，或者属性名同时撞上 row 属性和 table 列名时，会给出明确的选择方式，不再偷偷写错位置。
   Git 提交：`910c0fb5`, `769e7374`, `5799f66b`

4. **删除和移动操作更符合 Excel 行为。** 删除所有数据行后保留一个空数据行；隐藏最后一个可见 sheet 会被拒绝；切片器仍引用 Pivot 时不允许直接删除。
   Git 提交：`8cc3081c`, `5da3adaa`, `74dff81d`

### 3.4 Excel 图表、图片和绘图对象

1. **图表数据和显示更可靠。** waterfall 图的 legend、secondary axis 标题、图表缓存、scatter/bubble 引用和跨 sheet 图表引用都得到修复。
   Git 提交：`49af2898`, `b96b4b34`, `1663a241`, `df90233d`, `3c8c7423`

2. **图片和 shape 的位置、大小、裁剪、旋转和翻转更稳定。** 不会轻易写到工作表边界外，读回结果也更接近原始文件。
   Git 提交：`0cb0fc24`, `0f14fe82`, `1bd18ea3`, `5e212405`, `d8bf29dc`

3. **HTML 预览中的图表位置更准确。** 图表的 anchor 偏移会被计算进去，多个图表重叠时布局更接近 Excel。
   Git 提交：`cac644c0`, `4e3e91b4`

### 3.5 Pivot、Slicer、条件格式和数据验证

1. **Pivot dump 输出真实字段名和布局。** 不再输出难以使用的内部索引，compact/outline/tabular 布局也能正确识别。
   Git 提交：`23fcd207`, `83313c16`

2. **条件格式和 slicer dump 更完整。** 不再只输出第一条条件格式规则或第一个 slicer。
   Git 提交：`943b3917`, `8ce1c6f7`, `3c7a477b`

3. **数据验证和表格边界检查更严格。** 公式、范围、下拉框和错误样式的读写更符合 Excel。
   Git 提交：`9a5d451a`, `ebf1049f`, `5ea8b8c1`, `ca4f1e67`

### 3.6 查询、selector 和 batch

1. **复杂查询表达式更可靠。** 支持 `not(...)`，逗号 selector 做并集，`=`/`!=` 数字比较符合 Excel 数字语义。
   Git 提交：`33c851e0`, `04b4e0dc`, `199e564a`

2. **正则表达式错误不再静默变成 0 条结果。** 错误的正则、带特殊字符的值和 selector 中的 `!` 会得到明确诊断。
   Git 提交：`95761c0e`, `03d81d0b`, `21158912`

3. **row 查询更容易排错。** 未知列会列出可用列名；检测到的表结构不明确时会说明原因；0 条结果和真正的错误不再混在一起。
   Git 提交：`dc25f0e5`, `a2265a41`, `3f88a10a`, `8a8f091d`

4. **batch 参数更安全。** `query` 支持 `path` 别名，但缺少 selector 会按单项报错，不会误把空 selector 当成“匹配全部”。
   Git 提交：`06abe02f`

5. **未知命令参数不再被吞掉。** 写错 `--option` 时会立即报错并提示正确写法。
   Git 提交：`3a5e4772`

6. **查询输出和错误提示的建议算法统一。** 属性名拼写建议使用统一的 Damerau-Levenshtein 实现。
   Git 提交：`8b87da13`

### 3.7 Resident 和核心稳定性

1. **同一个文件只允许一个 resident。** 多个并发命令不会再各自打开一份内存副本，然后互相覆盖文件。
   Git 提交：`2285bfcf`

2. **空闲保存会给新命令让路。** Excel 公式扫描被命令打断后会在下一次空闲继续，避免用户一直等保存。
   Git 提交：`8062abc8`, `d3a52e8d`

3. **子进程通信更不容易死锁。** stdout/stderr 两条管道都被消费，并且等待时间有上限。
   Git 提交：`f7acea3a`

4. **浏览器截图遇到慢资源不再无限挂起。** headless Chrome 的外部资源加载有超时和降级路径。
   Git 提交：`b83334d0`

### 3.8 Word、PPTX 和渲染

1. **Word 复杂内容预览更完整。** 表格里的多级列表、页眉页脚列表、脚注尾注、VML 横线、altChunk、bookmark、content control 和字段内容得到修复。
   Git 提交：`bfccf47d`, `ffcc6fba`, `4f9555f7`, `fc90e44b`, `60c2a9ae`, `124440ad`

2. **Word 样式继承更准确。** basedOn 样式链中的边框、tab、cell margin、numId 和 ilvl 不再只继承一部分。
   Git 提交：`b81ae018`, `11ab4f54`

3. **公式渲染链路更稳定。** OMML 转 LaTeX 后使用 KaTeX，web font 通过缓存代理优先加载，CDN 作为兜底。
   Git 提交：`69a04774`, `4c208926`, `ba27cd33`, `f54d8f26`

4. **PPTX 文本自动适配尺寸更准确。** 按真正的文字区域计算，而不是拿整个 shape 外框计算。
   Git 提交：`4e1ed33f`

## 4. 部分合并：代码或辅助文件存在，但用户入口没有完整接通

1. **Schema CRC 辅助类存在，公开命令未接通。** `/Users/xsq/Desktop/supdone/projects/thatagata-agent/office-cli/OfficeCLI/src/officecli/Help/SchemaCrc.cs` 还在，但 `officecli --output-schema-crc` 的 `Program.cs` 早期分发没有合入。
   Git 提交：`7de674c2`, `96f9bacd`

2. **Excel 图表 helper 有不同实现。** main 增加了全局 chart part helper 和局部常量整理；当前 dev 使用了已有的 unanchored chart 处理方式，不是逐字合入 main 的实现。
   Git 提交：`ac727981`, `be608838`

## 5. 未合并：main 的扩展功能和配套代码

### 5.1 插件生态和 PDF 导出

1. **插件协议和插件发现机制未合并。** 包括 plugin manifest、进程管理、dump-reader、format-handler、插件目录扫描和 `officecli plugins` 命令。
   Git 提交：`2b32d4b3`, `6603d38f`, `62101e34`

2. **额外格式支持未合并。** main 可以通过插件把 `.doc` 等外部格式转换成 docx/xlsx/pptx，或直接通过 format-handler 提供 HTML/SVG/forms 视图。
   Git 提交：`2b32d4b3`

3. **PDF exporter 未合并。** main 增加过 `officecli export`，后来折叠为 `view <file> pdf`，并增加了解锁源文件后再导出的处理。
   Git 提交：`fa2bfd9f`, `b439d14e`, `1ceb8d42`

### 5.2 安装、自更新和 skills install

1. **`officecli install` 未合并。** main 支持把当前 binary 自安装到 `~/.local/bin` 或 Windows 用户目录。
   Git 提交：`1ed7b355`, `f1a1af1a`, `990cbc7d`

2. **自动安装 skill 未合并。** main 会检测 Claude、Cursor、Copilot、Codex、Pi、OpenCode 等工具，并自动复制 skill 文件。
   Git 提交：`46684d9e`, `ec8a332d`, `31e701e2`, `8be12f82`

3. **`skills list/install` 未合并。** main 支持列出 skill、安装全部 skill、安装单个 skill 和指定 agent 安装。
   Git 提交：`6ca4486d`, `dcd8c755`, `21920d00`, `ff6225ed`

4. **后台更新检查和自更新未合并。** main 会每天检查版本、下载 `.update` 文件、校验 SHA256、在下一次启动时替换旧 binary。
   Git 提交：`ab52656d`, `c0c7305b`, `cff1e846`, `f18bbae2`, `4cf610d8`

5. **MCP 保留，但 MCP 进程中的后台自动更新没有保留。** 当前分支继续提供 MCP server 和 MCP 注册命令，删除了 main 中定时自更新逻辑。
   Git 提交：`9dd637e6`, `c0950a46`, `c04c5281`

### 5.3 SDK 自动安装和外部安装链路

1. **Node SDK 未合入自动找 binary、自动下载和 `install()` API。**
   Git 提交：`1648b007`, `6ce96cd3`, `5040f927`

2. **Python SDK 未合入自动安装和 mirror-first 下载逻辑。**
   Git 提交：`186a0e5e`, `6ce96cd3`

3. **npm 安装器和相关发布文件未合并。** 包括 `npm/install.js`、平台 binary 下载、npm publish 和 SDK smoke 流程。
   Git 提交：`b1af162d`, `68d66669`, `65893f46`

### 5.4 CLI 辅助能力：已合入与未合入项

1. **`query --compact --fields` 已合并（待提交）。** 已提供稳定的 TSV 输出：一行一个元素、最后一行 total；`--fields` 可追加 Format 字段。CLI 入口、direct 模式和 resident 模式均已接通。
   Git 提交：`ae7fcd54`, `16c7c6ae`

2. **`officecli --output-schema-crc` 公开参数未合并。** 这是给 SDK/自动化工具判断 schema 是否变化的兼容性指纹。
   Git 提交：`7de674c2`, `96f9bacd`

3. **CLI 文件日志未合并。** main 支持 `officecli config log true`，把命令、输出和错误记录到 `~/.officecli/officecli.log`。
   Git 提交：`2c6bae41`

4. **`@` 属性参数的 response-file 兼容修复已合并（待提交）。** 已关闭 `@...` 的 response-file 自动替换，使 `--prop @height=25` 可以到达 selector 解析器。
   Git 提交：`769e7374`

5. **Mermaid 缓存的每日刷新未合并。** 这项能力依赖后台更新检查，只会刷新已经存在的缓存，不会首次主动下载。
   Git 提交：`2785590e`, `ab52656d`

### 5.5 文档、版本和测试配套

1. **main 的英文 README、CONTRIBUTING 和安装说明未合并。** 当前 dev 保留中文帮助和中文 skill，这是有意选择，不是遗漏。
   Git 提交：`6e3682de`, `b0502fe7`, `28d799e6`, `f7a21528`

2. **main 的 skill 文档结构和 TOC 文案更新未按英文版本合入。** 当前 dev 保留中文正文；功能代码和 schema 行为仍以当前 dev 为准。
   Git 提交：`1171d919`, `0fede920`, `4e06c139`, `f7a21528`

3. **main 的版本号 1.0.135 和对应 CHANGELOG 未合入。** 当前 dev 仍保持 `1.0.132`，避免直接覆盖当前分支的版本线。
   Git 提交：`d2d9c60f`, `5a73c41f`, `ffc118c7`

4. **main 历史中可恢复的测试已独立恢复。** `020fc7a1` 提供测试工程与 OLE/图片测试，`c08ebfc7` 提供 SDT 契约测试；它们后来分别在 `c59b4dee`、`ac7c0ce0` 被删除。现已放入 `tests/OfficeCli.Main.Tests`，与当前 `tests/OfficeCli.Tests` 的 146 个 dev 手工测试隔离；为适配已演进的产品语义，OLE 查询和 SDT checkbox 断言已同步更新。
   Git 提交：`020fc7a1`, `c08ebfc7`, `c59b4dee`, `ac7c0ce0`

5. **main 的安装 Smoke Test、平台 portability 检查和 npm CI 未合并。** 这些主要验证发布流程，不改变 OfficeCLI 的运行时能力。
   Git 提交：`65893f46`, `7f0e1482`, `68d66669`

## 6. 价值判断建议

### 建议继续保留

1. **Excel 输入校验、失败回滚、公式和表格关联修复。** 这些直接决定文件会不会损坏，属于核心能力。
   Git 提交：`e9197b87`, `dd065553`, `2b1aa731`, `dcdb6f64`

2. **查询 selector 的语义修复和错误提示。** AI 最怕“返回成功但结果错了”，这批修复能明显降低误操作。
   Git 提交：`33c851e0`, `95761c0e`, `06abe02f`, `3a5e4772`

3. **Resident 单例锁、自动保存让路和子进程管道修复。** 这些属于运行时基础设施，建议保留。
   Git 提交：`2285bfcf`, `8062abc8`, `f7acea3a`

4. **Word 预览、公式、列表、bookmark 和字段 round-trip。** 这些是 Word 主干场景的直接质量提升。
   Git 提交：`bfccf47d`, `c100244d`, `b81ae018`, `69a04774`

### 值得单独评估

1. **`query --compact --fields`（已合并）。** 适合 AI 查询大型 PPTX/Word；后续只需补充真实大文档与格式字段的回归测试，不需要再讨论是否恢复。
   Git 提交：`ae7fcd54`, `16c7c6ae`

2. **`view --range`（已合并）。** 对大型 Excel 和大页面截图很有价值；后续应补充区域截图与边界 range 的回归测试，不需要恢复插件生态。
   Git 提交：`f6272850`, `d9cfa304`, `7ebb33d8`

3. **Schema CRC。** 对 SDK、缓存和自动化兼容性检查有价值，对普通 CLI 用户不重要。
   Git 提交：`7de674c2`, `96f9bacd`

4. **双来源回归测试集。** main 历史测试与 dev 手工测试分为独立工程：前者用于保留 main 的可追溯覆盖，后者继续覆盖当前扩展能力；安装器、npm、平台 portability 等发布链路测试仍未恢复。
   Git 提交：`020fc7a1`, `c08ebfc7`, `c59b4dee`, `ac7c0ce0`

### 当前方向下建议继续不恢复

1. **插件生态、额外格式和 PDF exporter。** 除非产品明确要支持 `.doc`、PDF 或第三方格式，否则维护成本明显高于当前收益。
   Git 提交：`2b32d4b3`, `b439d14e`, `fa2bfd9f`

2. **自动安装、后台更新和 skills install。** 它们会改变用户机器、访问网络并引入权限和供应链问题，不属于 Office 文档处理核心。
   Git 提交：`1ed7b355`, `ab52656d`, `6ca4486d`

3. **SDK 自动下载 binary。** 方便新用户，但会让 SDK 从“调用已有 CLI”变成“负责安装和升级 CLI”。
   Git 提交：`1648b007`, `6ce96cd3`, `186a0e5e`
