# PPTX Agent Skill

Author: xiesq, 2026-07-15

## ⚠️ Help 优先规则

**本 Skill 说明如何制作高质量 slide，而非罗列所有命令选项。属性名、enum 值或别名不确定时，务必先查阅 help，不要猜测。**

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
      "shape"
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
      "add",
      "chart"
    ]
  }
}
```

Help 反映当前安装的 CLI 版本。Skill 与 help 不一致时，**以 help 为准**。遇到以下情况应立即运行 help：`UNSUPPORTED props:` 警告、未知 animation preset、`connector.shape=` enum 变化，或 prop 与 alias 的差异（`lineWidth` 与 `line.width`、`color` 与 `font.color`）。

`{{OFFICE_LOAD_SKILL_TOOL}}` 只读取指导，不安装或修改本地 skills。

## 宿主工具与执行规范

（对应原 CLI「Shell 与执行规范」；shell 引号改为 Tool JSON 传参。）

**Path 与特殊字符。** element path 含 `[]`，货币值可能含 `$`：

- path 原样传入 `command_arguments` / Batch `path`（含 `[N]`），无需 shell 引号。
- `$` 写在 JSON 字符串内即可（无 shell expansion）。货币写 `"$1.42"` / `"$15M"`，不要写成 shell 转义的 `\$`。
- props 中的 `\n` / `\t` 仍由 Office 工具解释，且在 pptx / docx / xlsx 中一致：`\n` 为换行/paragraph 换行，`\t` 为 tab；字面量反斜杠+n 写 `\\n`。
- Batch JSON 内 `\n` / `\t` 均可正常工作。

如有疑问，写入后 Run `view text` 并逐字符比对。

**增量执行。** 每条操作后检查结果，再继续。包含 50 条操作的流程若在第 3 条失败，后续会静默连锁失败。任何结构操作（新建 slide、chart、animation、connector）之后，先 Run `get`，再叠加更多操作。

**阶段 0 Batch/Run 契约（固定补丁，非业务加戏）。**

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。

示例 path 使用 `/workspace/deck.pptx`；必须替换为真实目标文件。

## 输出要求

这是每个 deck 都必须满足的交付标准。违反任意一项即视为未完成，与内容质量无关。

### 全部 deck

**每张 slide 只表达一个观点。** 如果一张 slide 需要第二个标题才能说明其内容，就应拆分。密集的“关于 X 的一切”slide 会在 3 秒内失去观众。使用 section divider 组织相关的单观点 slide，不要堆成巨型 slide。

**显式设置文字层级，不要依赖 theme 默认值。** Theme 默认值会在不同 master 之间漂移；每个文本 shape 都应明确设置字号。

| 元素 | 最小值 | 常用值 | 最小 shape 高度 |
|---|---|---|---|
| Slide 标题 | **≥ 36pt**，bold | 36–44pt | ≥ 2cm |
| Section / subtitle | ≥ 20pt | 20–24pt | ≥ 1.2cm |
| 正文 | **≥ 18pt** | 18–22pt | ≥ 1cm |
| Caption / axis label | ≥ 10pt，muted | 10–12pt | ≥ 0.6cm |

经验公式：**最小 shape 高度 ≈ font_pt × 0.05cm**。将 18pt sublabel 放入 0.8cm 高的框会溢出，`view annotated` 能发现这一问题。

标题必须**至少是正文的 2 倍**（36pt 配 20pt 可行；28pt 配 20pt 显得拘谨）。正文可低于 18pt 的四种合理例外：chart axis label、legend、footer / page number，以及不超过 5 个词的 KPI sublabel（如“活跃用户”）。描述性句子必须不小于 18pt。正文左对齐；仅标题与 hero number 居中。如果“card 放不下”，应减少 card，而非缩小 font。

**最多两种 font、一个 palette。** 使用一个 heading font 和一个 body font（例如 Georgia + Calibri）；只有大号数字或封面标题可使用第三种 *display* font，且不得破坏 heading+body 搭配。采用一种主导 brand color（占 60–70% 视觉权重）+ 一种辅助色 + 一种 accent；正文中不要混用 4 种以上颜色。**Design Principles 的 palette 与 font pairing 是下限，而非菜单：**用户提供 brand color/font 或既有 template 时，优先匹配它们；否则可将命名组合视为校准种子，自由混合或偏离，但结果不能更差，且仍须达到 contrast 下限。

**每张 slide 都应包含一个能传递信息的非文本 visual。** 可使用 shape、chart、icon 或 gradient band，但不能只是装饰。只含 bullet 的 deck 与 Word document 没有区别。例外是纯引用 slide、code block 与单张 summary table slide。

**少即是多，每个 element 都应有存在理由。** 上述 visual rule 是为了避免 bullet wall，并不意味着可以堆砌内容。不要以不传递信息的 decorative stat、icon 或 filler section（“data slop”）凑数。slide 看似空白时，应通过 layout 与 whitespace 改善，而非编造内容：宁可缩小范围并标注需要补充的内容，也不要擅自添加。

**每张 content slide 都需要 speaker note。** 使用 `type=notes` + `text=...`。speaker 需要讲稿，audience 不应逐字阅读 slide。

**文案应像人写的，而非 AI 生成。** title 应指向内容，而非追求 punchline。不要使用“不是 X，而是 Y”、人为制造 tension、伪洞察（如“magic moment”）或单词式戏剧感（如“Momentum.”）。删除 seamless、robust、game-changing 等夸张形容词，让数字本身传达信息。

**保留既有 template。** 文件已有 theme 和 master 时应与其匹配；既有约定优先于本指南。

### 视觉交付底线（适用于每个 deck）

宣布完成前，每张 slide 的 render（见 QA）必须满足：

- **不得将 placeholder token 渲染为内容。** chart 标题中不应出现 `{{name}}`、`$fy$24`、`<TODO>`、`lorem`、`xxxx` 或空的 `()`/`[]`。
- **不得有 shape 越过 slide 边界，也不得有被裁切的 text。** `view issues` 会标记两者（`shape_off_slide` 与 text-fit hint）。修复裁切时，增大 text box 或缩短内容，绝不能裁掉内容以强行放入。
- **cover 必须带有导向信息。** 应包含 title + subtitle + presenter/client + date + brand band 或 key-takeaway strap。只有 title 的 cover 看起来像未完成草稿。周围留出充分 whitespace 仍是正确做法；丰富不等于拥挤。
- **Contrast。** `view issues` 会自动标记常见情况：shape 自身 dark fill 上的 opaque dark text（`low_contrast`）。它无法识别 icon / chart-series fill、scheme/inherited color，或叠在**单独** background shape 上的 text。因此，对任何 brightness < 30% 的 fill（`1E2761`、`36454F`、deep forest / berry / cherry），仍应确认每个 body run、card body、chart series 与 icon 为 `FFFFFF` 或 brightness > 80%。mid-gray（`6B7B8D` ≈ 44%）在 laptop 上尚可读，但投影时会消失。处理 dark fill 后用 `view html` spot-check。

如果有任何失败，请在宣布完成之前停止并修复。

## 设计原则

deck 不是 document。观众只有 3 秒理解每张 slide。添加任何内容前先问：“如果观众只读最大的 element，再快速看一眼，能否理解要点？”若必须阅读 bullet 才能理解，说明最大的 element 设计错了。

### 网格、边距、负空间

标准宽屏为 **33.87 × 19.05cm**。在内部将其视为 12 列网格：

- **四边 margin ≥ 1.27cm**（0.5"）。
- **card / column / row 的 block gap ≥ 0.76cm**（0.3"）：选择一个值（0.76 或 1.27cm）并始终使用；混合 gap 看起来未完成。
- **每张 slide 至少 20% whitespace。** 填满每一个像素显得业余。
- **做构图，不要机械居中。** whitespace 是结构的一部分：内容偏上、下三分之一保留空白，是正确构图而非“空缺”。有意的不对称（左侧内容、右侧留白）往往比所有内容居中更具设计感；不要仅因存在空白就把它填满。
- 对于卡片网格：`usable = 33.87 − 2·margin − (N−1)·gap`，然后 `col_width = usable / N`。不要手动选择 x 坐标。

### Font pairings

按 document register 配对 font，而不是追求新奇。“Best For”仅是提示，不是强制规则；table 外的 pairing 只要适合也可使用。以下 8 种是 seed，不是固定集合。

| Heading | Body | 适用场景 |
|---|---|---|
| Georgia | Calibri | 正式 business、finance、executive report |
| Arial Black | Arial | 大胆 marketing、product launch |
| Calibri | Calibri Light | 简洁 corporate、minimal design |
| Cambria | Calibri | 传统专业、legal、academic |
| Trebuchet MS | Calibri | 亲和 tech、startup、SaaS |
| Impact | Arial | 大胆 headline、event deck、keynote |
| Palatino | Garamond | 优雅 editorial、luxury、nonprofit |
| Consolas | Calibri | developer tool、technical / engineering |

在每个 shape 上显式设置 fonts（标题上为 `font=Georgia`，正文上为 `font=Calibri`），而不是通过 theme 继承。

### 颜色和对比度

各 column 含义：**Primary**（主导色，占 60–70% visual weight，也是第一眼看到的 color）、**Secondary**（辅助色）、**Accent**（少量、单点强调）、**Text**（light fill 上的 body text）、**Muted**（caption / axis label / footer）。

| Theme | Primary | Secondary | Accent | Text | Muted |
|---|---|---|---|---|---|
| Coral Energy | `F96167` | `F9E795` | `2F3C7E` | `333333` | `8B7E6A` |
| Midnight Executive | `1E2761` | `CADCFC` | `FFFFFF` | `333333` | `8899BB` |
| Forest & Moss | `2C5F2D` | `97BC62` | `F5F5F5` | `2D2D2D` | `6B8E6B` |
| Charcoal Minimal | `36454F` | `F2F2F2` | `212121` | `333333` | `7A8A94` |
| Warm Terracotta | `B85042` | `E7E8D1` | `A7BEAE` | `3D2B2B` | `8C7B75` |
| Berry & Cream | `6D2E46` | `A26769` | `ECE2D0` | `3D2233` | `8C6B7A` |
| Ocean Gradient | `065A82` | `1C7293` | `21295C` | `2B3A4E` | `6B8FAA` |
| Teal Trust | `028090` | `00A896` | `02C39A` | `2D3B3B` | `5E8C8C` |
| Sage Calm | `84B59F` | `69A297` | `50808E` | `2D3D35` | `7A9488` |
| Cherry Bold | `990011` | `FCF6F5` | `2F3C7E` | `333333` | `8B6B6B` |

按主题选择，而不是默认选择：财务适合 Midnight Executive，产品发布适合 Coral Energy，安全/LOTO 适合 Cherry Bold。若最接近的命名 theme 不够贴切，可混合使用（例如 Forest & Moss 的 Primary + 金色 `D4A843` accent）。浅色填充使用 **Text**，标题/axis/footer 使用 **Muted**，深色填充上的正文使用 `FFFFFF` 或 Secondary。

dark background 上的 text 与 chart series 必须满足上述 contrast 底线。

### Chart choice 决策表

错误的 chart type 会直接破坏 3 秒理解测试：

| Data shape | 使用 | 避免 |
|---|---|---|
| category comparison（A vs B vs C） | `column`（vertical）/ `bar`（≥ 6 category 时 horizontal） | pie（slice 混在一起）、line（没有 time axis） |
| time series，1–3 个 series | `line` | area（occlusion）、bar（暗示 discrete） |
| part-of-whole，2–5 个 slice | `pie` / `doughnut` | 8+ slice 的 pie（不可读） |
| correlation / distribution | `scatter` | line（暗示有顺序） |
| multi-category × metric，密集 | stacked `column` 或 heatmap | 每个 metric 一个 chart，应合并 |
| KPI snapshot（单一大数字） | **large-text shape**（60–72pt + ≤ 5-word sublabel），不是 chart | gauge chart、tiny bar |

经验法则：如果 > 3 个系列且 > 8 个类别，则分成两个 charts 或切换到 table。

### Animation

animation 的使用量应服从 brand 与 content：正式 finance deck 通常接近于零，product launch 可以更具表现力。animation 是工具，不是装饰。以下三条底线能避免伤害 deck（不限制总使用量）：

- **有目的：**每个 animation 都应揭示或强调内容（progressive bullet reveal、cumulative chart），绝不只为装饰。无助理解就删除。
- **优雅降级：**pptx animation 在 Keynote / Slides / web / mobile 等 viewer 中 render 不一致，且可能完全不播放，因此每张 slide 作为**static frame**也必须正确可读。不要将必要信息隐藏在 reveal 后。
- **Live 验证：**animation 仅在 runtime 可见；`view html` 与 screenshot 看不到它，交付前必须在实际 presentation viewer 中确认。

风格建议（不是禁令）：使用快速 duration（约几百 ms）的 `fade` / `appear` / 单个 `zoom-entrance` 适合大多数 deck；`bounce` / `swivel` / `spin` / `fly-from-edge` / 密集 multi-object choreography 通常显得业余，只在 brand 有意表现 playful 时使用。

### Layout pattern 与 data display

slide 之间应改变 layout；重复相同 pattern 会让每张 slide 感觉一样。以下是常见 building block，而不是完整集合；每张 slide 选用一个，或按 content 需要在 table 外自定义 layout：

| Pattern | 使用场景 | 关键尺寸 |
|---|---|---|
| **Two-column**（左 text、右 visual） | concept + evidence；feature + screenshot | 每 column ≈ 14–15cm；gap 1cm |
| **Icon row**（填充 circle 内的 icon + bold header + description） | feature list、benefit、team role | icon circle 1.5–2cm；最多 3–4 row |
| **2×2 / 2×3 grid**（card tile） | quadrant analysis、SWOT、option comparison | gap ≥ 0.76cm；card 高度一致 |
| **Half-bleed image**（完整左/右半区，另一半叠放 content） | hero moment、case-study opener | image 宽 16–17cm；content column ≥ 14cm |
| **Large stat callout**（60–72pt number + 下方 ≤ 5-word sublabel） | single KPI、milestone、market size | 用 shape，不用 chart；sublabel 14–16pt muted |

**Data display 快速规则：**

- 2–3 个 option 的 before/after 或 A vs B comparison，用 comparison column 胜过 table。
- timeline 与 process flow 使用 numbered-step shape + connector，不要使用 bullet list。

### Image treatment（仅当 slide 使用 photo / screenshot / logo 时）

**先读取 image**（打开文件），再基于实际画面选择 treatment；不要只凭文件名盲目放置。

- **Full-bleed photo** → 以 COVER 方式铺满区域（裁切边缘），无 border。
- **Screenshot / diagram / logo** → 以 FIT 方式放置（不得裁切 content）。透明或 FIT image 应置于 contrasting fill 上，在其后放入彩色 rectangle，不能漂在白色上。
- **Text over photo** → 绝不要直接叠在 image 上。应放入 card，或在 image 与 text 之间加 protective scrim（约 50–60% opacity 的 dark rectangle，或从 text 边缘渐隐的 gradient）。
- 不得 stretch（扭曲 aspect ratio），不要在繁杂 screenshot 上覆盖 text。
- 优先使用用户提供的 image / brand asset；除非明确要求，不使用 emoji 或自绘艺术。

### Visual motif 承诺

选择一个独特 element（rounded image frame、filled circle 内的 section number、single-side border band、diagonal accent strip），并贯穿整个 deck；只设计一张 slide 而让其余 slide 保持普通，会显得像半途放弃。secondary motif 仅在不与 primary motif 竞争时使用。先在 build plan 中声明，例如：`## Motif: numbered circles in brand color`。

### 应避免的视觉 AI 痕迹

- **slide title 下不要使用装饰性 underline。** heading 下方的 stripe / rule 是最常见的 AI-slide 痕迹，应改用 whitespace 或 background-color change。
- **不要使用带彩色 left-border accent stripe 的 rounded-corner card。** 这是另一种典型 AI-slide 痕迹；改用 solid fill、top accent band 或 whitespace separation。
- **不要用 emoji 充当 icon**，除非 brand 本来使用它们；请使用 shape 或真实 icon asset。

文案层面的 AI 痕迹见“文案应像人写的，而非 AI 生成”。

## 通用工作流程

1. **打开/保存生命周期。** 开始时 Run `open`，结束时 Run `save`，将编辑 flush 到磁盘。`save` 只写入，保留 resident 以供后续编辑；仅在需要立即释放 resident（一次性交接）时 Run `close`。两者始终安全，不会报错或丢失工作。重复的 shape 网格使用 Batch。**只在非 Office 工具边界 flush：**宿主自身读取始终能看到编辑；仅在非 Office 工具之外的程序（python-pptx、PowerPoint、renderer、交付流程）读取文件前运行 `save`/`close`。
2. **先确认现状。** 新建 deck：Run `create`。已有 deck：先 Run `view outline`。切勿盲目编辑。
3. **先编排 title sequence（只计划，暂不构建）。** 创建任何 slide 或 shape 前，写出完整的、有序 slide title list。若只读 title 的人无法理解论点，应立刻修复 narrative arc；此时修改远比完成 14 张 slide 后便宜。选择一种 title grammar：全是 topic noun phrase，或全是 action statement，绝不混用，并贯穿全程（见“文案应像人写的，而非 AI 生成”）。
4. **按展示顺序构建。** 按观众观看顺序添加 slide：cover → agenda → section 1 divider → section 1 content → section 2 divider → … → close。`slide add` 的 `index` 虽可用，但线性 append 更利于可读性，也避免 index arithmetic error。**最终交付前，确认 slide count 与 narrative arc 匹配 build plan。** Gate 3 的 order-sanity check 可发现 cover 意外成为 14 张中的第 11 张，而不是第 1 张。
5. **逐张构建。** 先创建 slide + background，再添加 title，最后添加 supporting shape / chart / connector。定制设计始终使用 `layout=blank`。每次 structure operation 后，执行 `get /slide[N] depth 1` 确认 shape ID。
6. **按规范 format。** 根据要求添加 table；formatting 是交付内容，不是润色。
7. **保存并验证。** `save` 会 flush 到磁盘（或 `close` flush 并结束 session）。交付前务必在 target presentation viewer 中打开，因为 chart color、animation、font 与 zoom 都是 `view html` 无法 render 的 runtime feature。下方 QA 会进行完整验证。
8. **QA — 假设存在问题。** 修复并验证，直到一个周期发现零个新问题。

## 快速入门

最小可行 deck：cover + 一张 content slide + notes。请按实际文件名调整。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": [
      "/workspace/deck.pptx"
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
      "/workspace/deck.pptx"
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
      "/workspace/deck.pptx",
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
      "/workspace/deck.pptx",
      "/slide[1]",
      "--type",
      "shape",
      "--prop",
      "text=FY26 Strategic Review",
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
      "/workspace/deck.pptx",
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
      "/workspace/deck.pptx",
      "/slide[2]",
      "--type",
      "shape",
      "--prop",
      "text=Revenue grew 18% YoY",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1.2cm",
      "--prop",
      "width=30cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761"
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
      "/workspace/deck.pptx",
      "/slide[2]",
      "--type",
      "shape",
      "--prop",
      "text=Enterprise renewals + new EMEA region drove the beat; NRR held at 118%.",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=30cm",
      "--prop",
      "height=3cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=20",
      "--prop",
      "color=333333"
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
      "/workspace/deck.pptx",
      "/slide[2]",
      "--type",
      "notes",
      "--prop",
      "text=Lead with the 18% beat, preview EMEA."
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
      "/workspace/deck.pptx"
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
      "/workspace/deck.pptx"
    ]
  }
}
```

基本流程：open → slide + background → title → body → notes → save → validate。

## 阅读与分析

先宽后窄。先用 `outline` 了解整体，再使用 `view text` / `get` / `query` 深入查看。

Run `view`：`outline`（slide count + titles）→ `annotated` → `text start 1 end 5` → `issues` → `stats`（含 pictures missing alt）。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": [
      "/workspace/deck.pptx",
      "outline"
    ]
  }
}
```

**检查单个 element。** 使用 XPath-style path，索引从 1 开始，原样传入 JSON。优先使用 `@name=` / `@id=` selector，而不是位置 `[N]`，前者在 reorder 后更稳定。`[last()]` 有效。需要 machine-readable output 时加 `json`。示例：`/slide[1]`（`depth 1`）、`/slide[1]/shape[@name=Title]`、`/slide[1]/table[1]`（`depth 3`）。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[1]",
      "--depth",
      "1"
    ]
  }
}
```

**跨 deck 查询。** CSS-like selector；operator 包括 `=`、`!=`、`~=`、`>=`、`<=`、`[attr]`、`:contains()`、`:no-alt`。见 Help `query`。示例：`shape:contains("Revenue")`、`picture:no-alt`、`shape[fill=1E2761]`、`shape[width>=10cm]`。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "query",
    "command_arguments": [
      "/workspace/deck.pptx",
      "picture:no-alt"
    ]
  }
}
```

**`query json` output shape。** 结果位于 `.data.results[]`（例如 `.data.results[0].format.id`），不是 `.[0].id`。shape name 位于 `.name`，fill 位于 `.format.fill`，text color 位于 `.format.textColor`。用返回结果判断，不要用 shell `jq`。

**Visual preview（LEAD）。** Run `view html`（读取返回的 HTML path 做 per-slide visual audit）；单页可用 `view svg start 3 end 3`（charts + gradients 在 SVG 中不 render）。

**读取 output 时的预期行为：**

- **`layout=blank` 没有 title placeholder。** title 是普通 `shape` element，因此 `view outline` 显示 `(untitled)` 是**预期**，不是 defect。仅当 screen-reader outline compatibility 很重要时，才使用 `layout=title` + `placeholder[title]`。

## 创建和编辑

可用 verb：`add` / `set` / `remove` / `move` / `swap` / Batch / `raw-set`（仅 Run + **approval**）。一次 deck 构建的 90% 是 slide、shape、text、少量 chart、image 与 connector。

### Slide 与 background

slide 的 path 为 `/slide[N]`。定制设计始终使用 `layout=blank`。background 支持 solid color、gradient 或 image。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/deck.pptx",
    "operations": [
      {"command":"add","parent":"/","type":"slide","props":{"layout":"blank","background":"1E2761"}},
      {"command":"add","parent":"/","type":"slide","props":{"layout":"blank","background":"1E2761-CADCFC-180"}},
      {"command":"add","parent":"/","type":"slide","props":{"layout":"blank","background":"image:/path/to/hero.jpg"}}
    ],
    "stop_on_error": true
  }
}
```

### Shapes

`shape` 保存文本、填充、边框、位置和可选的 animation/链接。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[2]",
      "--type",
      "shape",
      "--prop",
      "name=Title",
      "--prop",
      "text=Key Insight",
      "--prop",
      "x=2cm",
      "--prop",
      "y=2cm",
      "--prop",
      "width=20cm",
      "--prop",
      "height=3cm",
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

position 是显式指定的：没有 layout engine，必须自行处理 grid math。`preset=` 选择几何 shape（`rect`、`roundRect`、`ellipse`、`triangle`、`arrow`、`star5` 等）；不支持自定义 `M...Z` path，应选择 preset。**创建时为 shape 命名**（`name=HeroTitle`），之后通过 `"/slide[N]/shape[@name=HeroTitle]"` 访问。name 在 z-order 改动或 delete-then-add 后仍可用，位置 `/shape[3]`（甚至 `@id=`）则会变化。每次 structure change 后若必须使用 positional index，应先重新 `get depth 1`。

### Shape 内的 text（paragraph、run 与 style）

shape 内含 paragraph（`paragraph[K]`）和 run（`run[K]`）。单行 text 可直接在 shape 上设置 `text=`；text 中的 `\n` 产生 paragraph break，`\t` 产生 tab。`add type paragraph` 支持与 shape 相同的 style prop（text、alignment、bold、italic、size、color、font）。同一行内需要 mixed style 时，追加带 style 的 run：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[2]/shape[@name=Card1]/paragraph[1]",
      "--type",
      "run",
      "--prop",
      "text= (inline detail)",
      "--prop",
      "size=14",
      "--prop",
      "italic=true",
      "--prop",
      "color=8899BB"
    ]
  }
}
```

### Chart

按“设计原则”中的 chart choice table 选择 chart type。完整 prop list（chart type enum、`seriesN.*`、`data=` / `categories=`、axis option）见 Help `add chart`。典型的 brand-color multi-series chart：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[3]",
      "--type",
      "chart",
      "--prop",
      "chartType=column",
      "--prop",
      "series1.name=Revenue",
      "--prop",
      "series1.values=42,45,48",
      "--prop",
      "series1.color=1E2761",
      "--prop",
      "series2.name=Growth",
      "--prop",
      "series2.values=2,7,7",
      "--prop",
      "series2.color=CADCFC",
      "--prop",
      "categories=Q1,Q2,Q3",
      "--prop",
      "x=2cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=20cm",
      "--prop",
      "height=10cm"
    ]
  }
}
```

注意：(1) 含 `()`、`[]`、`TBD` 的 chart title 会按字面量 render；(2) 部分 viewer 会将 chart color 标准化为 theme default，应在 target viewer 中验证。series 可在创建后添加（`add type series`）。

### Pictures

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[4]",
      "--type",
      "picture",
      "--prop",
      "src=hero.jpg",
      "--prop",
      "x=1cm",
      "--prop",
      "y=1cm",
      "--prop",
      "width=32cm",
      "--prop",
      "height=18cm",
      "--prop",
      "alt=Product hero, gradient lit from right"
    ]
  }
}
```

使用 `query 'picture:no-alt'` 确认；交付前结果必须为空。

### Connector（LEAD：flowchart / decision tree 的 first-class element）

在两个 shape 或自由坐标之间绘制 line。完整 prop / enum reference（`shape`、`headEnd` / `tailEnd` value、`from` / `to` reference format）见 Help `add connector`。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[5]",
      "--type",
      "connector",
      "--prop",
      "from=/slide[5]/shape[@name=BoxA]",
      "--prop",
      "to=/slide[5]/shape[@name=BoxB]",
      "--prop",
      "shape=elbow",
      "--prop",
      "color=333333",
      "--prop",
      "tailEnd=triangle"
    ]
  }
}
```

**每个 flow connector 都需要 arrowhead。** 没有 arrowhead 时，`bentConnector3` 会 render 成没有方向的 line。`preset=rightArrow` overlay 只适用于 horizontal flow；带发散 edge 的 diamond / decision tree 需要 `tailEnd=`。

### Animation（LEAD）

遵循上述 animation 底线（有目的、优雅降级、live 验证）。preset name 与 duration syntax 见 Help `animation`。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[2]/shape[@name=HeroCard]",
      "--prop",
      "animation=fade-entrance-400"
    ]
  }
}
```

### Hyperlink、tooltip 与 slide jump

`link=slide[N]` 用于 deck 内跳转（索引从 1 开始，target slide 必须存在）；`link=nextslide` / `firstslide` / `lastslide` / `previousslide` / `endshow` 用于 named navigation；`link=https://...` 用于 URL；`tooltip="..."` 用于 hover text。

### Table、placeholder、group、zoom：简表

- **Table：**`type=table` + `rows=N` + `cols=M`。row-level `set` 支持 `height` 与 `c1/c2/c3`（seed cell text）。header-row style 位于 table-level（`firstRow=true` / `headerFill=`），不是 row prop。cell format 设置在 cell paragraph / run 上。应先填充 row，再设置 table-level font，因为 row operation 会重置 font cascade。
- **Placeholder：**`"/slide[N]/placeholder[title]"` / `placeholder[body]`。只在 slide 使用带 placeholder 的 layout 时可用，`layout=blank` 不可用。
- **Group（LEAD）：**通过 `"/slide[N]/group[@name=G]/shape[1]"` 访问 child，比 positional index 更能抵御 reorder。
- **Zoom slide（LEAD）：**`type=zoom` + `target=N`（每个 target 一个 link，alias 为 `slide`）。multi-target navigation hub 需生成 N 个单独 zoom shape。zoom 是 runtime feature：`view html` 只显示 static geometry，zoom interaction 只在 live presentation viewer 中运行。
- **Slide comment：**reviewer annotation 锚定在 `/slide[N]/comment[M]`。完整 lifecycle 为 `add / set / get / query / remove`。prop 包括 `text`、`author`、`initials`（自动派生）、`date`（ISO 8601，默认 UtcNow）、`x` / `y`（length anchor）。

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[2]",
      "--type",
      "comment",
      "--prop",
      "author=Alice",
      "--prop",
      "text=Tighten this bullet",
      "--prop",
      "x=20cm",
      "--prop",
      "y=3cm"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "remove",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[2]/comment[1]"
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
      "/workspace/deck.pptx",
      "/slide[2]/shape[@name=HeroCard]",
      "--prop",
      "animation=none"
    ]
  }
}
```

计数 review comments：Run `query comment`（可加 `json`），用 `.data.results` 长度判断。

### Deck-level recipe

以下 pattern 仅从 primitive 难以推导。每个 recipe 先说明**visual outcome**，再给出可运行 block。使用 `/slide[last()]` 访问刚添加的 slide。这些 recipe 展示**structure 与 coordinate math**；应替换为本主题选择的 palette / font。navy `1E2761` + Georgia 只是示例 theme，不是要逐字复制的 house style。

**Z-order。** 后添加的 shape 位于上层。先添加 background decoration，最后添加 title。事后修复使用 `zorder=back/front`；此操作会重新编号 sibling，因此继续叠加前先重新 `get depth 1`。

#### (a) Cover（及 section divider）

**Visual outcome。** 深 navy fill，居中 44pt title，18pt ice-blue metadata line。

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/deck.pptx",
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
      "/workspace/deck.pptx",
      "/slide[last()]",
      "--type",
      "shape",
      "--prop",
      "text=Strategic Growth Review",
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
      "/workspace/deck.pptx",
      "/slide[last()]",
      "--type",
      "shape",
      "--prop",
      "text=Prepared for Acme Leadership — FY26 Outlook",
      "--prop",
      "x=2cm",
      "--prop",
      "y=11cm",
      "--prop",
      "width=29.87cm",
      "--prop",
      "height=1.2cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=18",
      "--prop",
      "color=CADCFC",
      "--prop",
      "align=center"
    ]
  }
}
```

**Section divider** = 相同 cover，加一个最先添加的大号半透明 number（`size=120`、`opacity=0.15`），使其位于 section title 后方。

#### (b) Data slide（chart + commentary card）

**Visual outcome。** 左侧 2/3：使用 brand-series color 的 column chart。右侧 1/3："Key Insight" card，20pt heading + 18pt body；观众先读取要点，再解析 bar。

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/deck.pptx",
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
      "/workspace/deck.pptx",
      "/slide[last()]",
      "--type",
      "shape",
      "--prop",
      "text=FY26 Revenue Beat Plan by 18%",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1cm",
      "--prop",
      "width=30cm",
      "--prop",
      "height=1.8cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=36",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761"
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
      "/workspace/deck.pptx",
      "/slide[last()]",
      "--type",
      "chart",
      "--prop",
      "chartType=column",
      "--prop",
      "series1.name=Actual",
      "--prop",
      "series1.values=42,45,48,55",
      "--prop",
      "series1.color=1E2761",
      "--prop",
      "series2.name=Plan",
      "--prop",
      "series2.values=40,42,45,48",
      "--prop",
      "series2.color=CADCFC",
      "--prop",
      "categories=Q1,Q2,Q3,Q4",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=3.5cm",
      "--prop",
      "width=20cm",
      "--prop",
      "height=14cm",
      "--prop",
      "title=FY26 Revenue ($M)"
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
      "/workspace/deck.pptx",
      "/slide[last()]",
      "--type",
      "shape",
      "--prop",
      "preset=roundRect",
      "--prop",
      "fill=F5F7FA",
      "--prop",
      "line=none",
      "--prop",
      "x=22.5cm",
      "--prop",
      "y=3.5cm",
      "--prop",
      "width=9.8cm",
      "--prop",
      "height=14cm"
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
      "/workspace/deck.pptx",
      "/slide[last()]",
      "--type",
      "shape",
      "--prop",
      "text=Key Insight",
      "--prop",
      "x=23cm",
      "--prop",
      "y=4cm",
      "--prop",
      "width=9cm",
      "--prop",
      "height=1.2cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=20",
      "--prop",
      "bold=true",
      "--prop",
      "color=1E2761"
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
      "/workspace/deck.pptx",
      "/slide[last()]",
      "--type",
      "shape",
      "--prop",
      "text=EMEA launch + NRR at 118% drove 12pp of the 18pp beat.",
      "--prop",
      "x=23cm",
      "--prop",
      "y=5.5cm",
      "--prop",
      "width=9cm",
      "--prop",
      "height=11cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=18",
      "--prop",
      "color=333333"
    ]
  }
}
```

#### (c) Flowchart / process diagram（box + connector）

**Visual outcome。** 四个 rounded box，y=8cm，每个 6×3cm，navy / ice-blue 交替，通过带 triangle arrowhead 的 elbow connector 连接。

grid math（4 个 box、33.87cm slide、1.5cm margin）：`gap = (33.87 − 3 − 24) / 3 = 2.29cm`。x position：`1.5, 9.79, 18.08, 26.37`。

每个 box 通过 `valign=middle` 承载自己的 label，不需要额外 overlay shape。使用 Batch 处理坐标。将 `$SLIDE` 替换为实际 slide 索引（如 `3`）。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/deck.pptx",
    "operations": [
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"Step1","preset":"roundRect","fill":"1E2761","line":"none","x":"1.5cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 1","font":"Georgia","size":"20","bold":true,"color":"FFFFFF","align":"center","valign":"middle"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"Step2","preset":"roundRect","fill":"CADCFC","line":"none","x":"9.79cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 2","font":"Georgia","size":"20","bold":true,"color":"1E2761","align":"center","valign":"middle"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"Step3","preset":"roundRect","fill":"1E2761","line":"none","x":"18.08cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 3","font":"Georgia","size":"20","bold":true,"color":"FFFFFF","align":"center","valign":"middle"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"Step4","preset":"roundRect","fill":"CADCFC","line":"none","x":"26.37cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 4","font":"Georgia","size":"20","bold":true,"color":"1E2761","align":"center","valign":"middle"}}
    ],
    "stop_on_error": true
  }
}
```

Connector pattern — 先用独立 Run readback 确认上一批 shape 全部成功且稳定 name 可寻址；之后三条 connector 彼此独立，可作为一个 Batch：

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/deck.pptx",
    "operations": [
      {"command":"add","parent":"/slide[3]","type":"connector","props":{"from":"/slide[3]/shape[@name=Step1]","to":"/slide[3]/shape[@name=Step2]","shape":"elbow","color":"333333","tailEnd":"triangle"}},
      {"command":"add","parent":"/slide[3]","type":"connector","props":{"from":"/slide[3]/shape[@name=Step2]","to":"/slide[3]/shape[@name=Step3]","shape":"elbow","color":"333333","tailEnd":"triangle"}},
      {"command":"add","parent":"/slide[3]","type":"connector","props":{"from":"/slide[3]/shape[@name=Step3]","to":"/slide[3]/shape[@name=Step4]","shape":"elbow","color":"333333","tailEnd":"triangle"}}
    ],
    "stop_on_error": true
  }
}
```

`shape=elbow` 是规范的（`bentConnector2` / `bentConnector3` 也被接受）。

#### (d) Multi-slide deck skeleton

没有 code block；这是 layout rhythm。以下 sequence 仅是**一种有效节奏的示例**（dark divider 与 white content 交替），不是必须的运行顺序。应先按内容确定实际 narrative arc（见“先编排 title sequence”），再采用适合的 divider/content rhythm：

- **10-slide review：**Cover · Agenda · 3 KPI · Div01 · Chart · Chart · Div02 · Flow · Timeline · Close
- **20-slide pitch：**上述 rhythm × 2，分为 Problem · Solution · Market · Product · Traction · Model · Team · Financials · Ask
- 每个 divider 必须位于对应 section content **之前**（Gate 3 order sanity）。
- cover/divider = (a)；chart page = (b)；process page = (c)；KPI page = (e)；decision page = (f)。

#### (e) KPI callout：giant-number card grid

**Visual outcome。** 一行中放置三个或四个 giant number；每张 card = unit sublabel + small percent-change chip + one-line takeaway。这是最常见的 executive-deck element。

**Size rule。** 60pt bold Georgia 在 9.78cm card 中可容纳约 5 个字符（`$84.2`、`118%`、`24.5`）。更长 value（`$84.2M`）应拆分为 `$84.2` 的 large number 与 `USD millions` 的 sublabel；不要为了塞入 unit suffix 而缩小 font，只会导致换行。

grid math（3 张 card、1.5cm margin、0.76cm gap）：`col_width = (33.87 − 3 − 1.52) / 3 = 9.78cm`。x position：`1.5, 12.04, 22.58`。只对一张 “watch” card 使用 accent color，使风险在一秒内可识别。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/deck.pptx",
    "operations": [
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"preset":"roundRect","fill":"1E2761","line":"none","x":"1.5cm","y":"4cm","width":"9.78cm","height":"7cm"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"text":"84.2","x":"1.5cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":true,"color":"FFFFFF","align":"center"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"text":"USD millions · ARR","x":"1.5cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"CADCFC","align":"center"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"text":"+24% YoY","x":"1.5cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","bold":true,"color":"CADCFC","align":"center"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"preset":"roundRect","fill":"B85042","line":"none","x":"22.58cm","y":"4cm","width":"9.78cm","height":"7cm"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"text":"$1.42","x":"22.58cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":true,"color":"FFFFFF","align":"center"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"text":"CAC payback (yrs)","x":"22.58cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"FFFFFF","align":"center"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"text":"+8% — watch","x":"22.58cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","bold":true,"color":"FFFFFF","align":"center"}}
    ],
    "stop_on_error": true
  }
}
```

#### (f) Decision tree：YES / NO branching

**Visual outcome。** top-center diamond；YES / NO child box 向左右分叉，再汇聚到 shared terminal box。layout：diamond `x=13.94, y=2cm, 6×3cm`；YES 在 `3cm, 7.5cm`；NO 在 `22.87cm, 7.5cm`；terminal 在 `13.94cm, 13cm`。约定：red = stop/escalate，blue = standard，green = safe terminal。**每个 connector 都需要 arrowhead**，否则读者会误判方向。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/deck.pptx",
    "operations": [
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"Decide","preset":"diamond","fill":"1E2761","line":"none","x":"13.94cm","y":"2cm","width":"6cm","height":"3cm","text":"Hazardous energy present?","font":"Calibri","size":"14","bold":true,"color":"FFFFFF","align":"center","valign":"middle"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"YesBox","preset":"roundRect","fill":"B85042","line":"none","x":"3cm","y":"7.5cm","width":"8cm","height":"3cm","text":"Lockout + Tagout + Verify","font":"Calibri","size":"16","bold":true,"color":"FFFFFF","align":"center","valign":"middle"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"NoBox","preset":"roundRect","fill":"CADCFC","line":"none","x":"22.87cm","y":"7.5cm","width":"8cm","height":"3cm","text":"Proceed with standard PPE","font":"Calibri","size":"16","bold":true,"color":"1E2761","align":"center","valign":"middle"}},
      {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"Done","preset":"roundRect","fill":"2C5F2D","line":"none","x":"13.94cm","y":"13cm","width":"6cm","height":"2.5cm","text":"Begin service","font":"Calibri","size":"16","bold":true,"color":"FFFFFF","align":"center","valign":"middle"}}
    ],
    "stop_on_error": true
  }
}
```

然后按 (c) 的 connector pattern 生成 4 个 connector（`Decide→YesBox`、`Decide→NoBox`、`YesBox→Done`、`NoBox→Done`）。

## QA（必须执行）

**假定存在问题。**第一次 render 几乎不会完全正确。如果发现零问题，说明检查还不够仔细。

### Delivery Gate（任一失败即 REJECT）

Gate 1–2b 属于 text/schema 层，无法看到 render 后的 slide；Gate 3 是唯一的 visual check。完成条件是所有 gate PASS，**且** Gate 3 loop 已收敛。

每个 gate 都是**独立 Run 一条 command，判断其 output**（不要 bash / grep / jq）。判断责任在执行者。

- **Gate 1 — schema。** Run `validate`。任何 schema error → REJECT 并修复。
- **Gate 2 — overflow / format / structure。** Run `view issues`。若列出*任一*问题（如含 `[O1]`、`[C1]`、`[S1]` 的行）→ REJECT、修复、重新运行直至干净。
- **Gate 2b — leftover placeholder。** Run `view text`，扫描 `xxxx`、`lorem` / `ipsum`、`<TODO>`、`placeholder`、"this slide layout" 与空 `()` / `[]`。任一命中 → REJECT。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": [
      "/workspace/deck.pptx"
    ]
  }
}
```

Gate 2 / 2b 同理：Run `view issues`、`view text`，按上文规则判 REJECT。

### Gate 3 — visual audit（强制）

选择**一种** path：

**Screenshot（默认）**：适用于有视觉能力的 agent。依次截图每张 slide：Run `view screenshot page 1 -o slide1.png`，然后 `page 2`，直到 page index 超出 deck（一个 screenshot = 一张 slide）。第 1 页即报错时，使用下方 fallback。

**以对抗性标准检查每个 PNG。**原则是“假定存在问题；发现不了问题，意味着检查不够仔细”。每个问题输出一行 `slide N: <issue>`，或输出 `PASS`。无论执行方式如何，这一步都必不可少。若环境支持 subagent，可交给一个*全新且独立*的 agent 判断：构建 deck 的 agent 天然偏向“看起来不错”，独立视角更关键。向其提供 screenshot、此 checklist 与相同的对抗性框架。没有 subagent 时，自己用同样标准检查。

**Fallback — HTML text**（无视觉能力或 screenshot 失败）：将 `view html` 作为 text 读取。DOM 无法证明**dark-on-dark / fine overlap / arrowhead / gap-margin metric / column alignment**，这些项应标记为“未做视觉验证”，而非 PASS。

**可选 `grid N`**：仅当用户需要检查 layout rhythm，或 `view outline` 显示异常 layout distribution 时使用：`view screenshot grid 3 -o grid.png`。

**每张 slide checklist（假定存在问题）：**

- **overlap**：shape / chart / giant decorative number（01/02/03 100pt+）相撞。
- **text overflow**：在 slide 或 shape 边界裁切，常见于 KPI card 与窄 box。
- **narrow text box**：内容虽技术上放得下，但换成许多短行（每行 1–2 个词）；例如 3cm KPI card 中的长 sublabel、过窄 column 中的 body line。
- **dark-on-dark**：fill brightness < 30%，而 text/icon brightness < 80%（包括没有 contrasting circle 的 dark icon on dark）。
- **image treatment**：photo 被拉伸/扭曲、text 直接叠在繁杂 image 上（无 card/scrim）、screenshot 或 logo 被裁切、透明 image 漂在白色上。
- **missing arrowhead**：flowchart connector 只是普通 line。
- **decorative-line / title mismatch**：accent bar 适合单行 title，但 title 换成两行，或反之。
- **footer / citation collision**：source line、page number 或 footnote 碰到上方内容。
- **tight margin / gap**：element 距 slide edge 约 0.5"，或两张 card 间距约 0.3"。
- **uneven gap**：一侧大面积空白、另一侧拥挤，导致 rhythm 破碎。
- **column / repeated-element misalignment**：KPI card / icon 偏离 baseline 或宽度不一致。
- **order sanity**：顺序符合叙事（cover → agenda → divider-before-section → close）。

若存在问题，输出 `slide N: <issue>` 并 REJECT；否则输出 `Gate 3 PASS`。HTML text fallback 还应添加 `<unverified-items> not visually verified`。

**Fix-verify（强制，最多 3 个 cycle）。** 修复 → 重跑 Gate 3 → 重复，直至没有新问题；一次修复经常暴露另一个问题。3 轮仍未收敛时，**停止**：原因可能是反复拉锯、template-level 原因或 agent 误判。报告 `slide N: <issue> — attempted: <fixes> — likely root: <template|design-conflict|ambiguous>`，并交由用户决定。

**最后 flush（Gate 的一部分）。** Gate 3 收敛后，以 Run `save` 收尾，确保编辑在交付前写入磁盘；若要一次性交接并释放 resident，可改用 Run `close`。这是必需的最后一步，不是可选项；两者都安全，不会报错或丢失工作。

## 常见陷阱

这是一份常见首次失败点的 checklist，涵盖 design 与传参陷阱。

| 陷阱 | 正确做法 |
|---|---|
| path 中的 `[N]` | 原样传入 JSON：`/slide[1]`（无需 shell 引号） |
| `name "foo"` | 所有 property 均通过 prop：`name=foo` |
| `/shape[myname]`（裸 name） | 使用 `@name=` selector：`/shape[@name=myname]` 或 `/shape[@id=10007]` |
| path 从 1 开始，而 `index` / Batch `index` 从 0 开始 | `/slide[1]` = 第一张 slide；`index=0` = 第一个位置 |
| `text=` 中的 `$` | JSON 字符串直接写 `"$15M"`（无 shell expand） |
| `text=` 中的 `\n` / `\t` | Office 工具会解释：`\n` = paragraph break，`\t` = tab。双 `\\n` 才表示字面量 |
