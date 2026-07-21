# Morph PPT Agent Skill

Author: xiesq, 2026-07-15

## ⚠️ Help 优先规则

**本 Skill 说明 Morph 工作流——何时同名配对、何时 ghost、CLI/宿主如何处理 transition——而不是罗列全部属性。** 不确定的 prop / enum / preset，先查 Help，禁止猜测。

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "pptx",
      "slide"
    ]
  }
}
```

再查 shape / animation / transition：

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
      "animation"
    ]
  }
}
```

Help 与已安装版本一致。Skill 与 Help 冲突时，**以 Help 为准**。确认 `transition=morph`、`advanceTime` / `advanceClick`、以及 `transition=morph-slow` / `morph-fast` / `morph-<DUR_MS>` 等 shorthand 是否在当前版本列出。

`{{OFFICE_LOAD_SKILL_TOOL}}` 只读取指导，不安装或修改本地 skills。先加载基座 `pptx`：

```json
{
  "tool": "{{OFFICE_LOAD_SKILL_TOOL}}",
  "arguments": {
    "command_arguments": [
      "pptx"
    ]
  }
}
```

## 心智模型与继承

**本 Skill 是 `pptx` 之上的 scene layer。** 基座硬规则全部继承、不重教：视觉交付底线（title ≥ 36pt / body ≥ 18pt / title ≥ 2× body）、33.87×19.05cm 十二栏网格、canonical palette、chart-choice、connector canon、Delivery Gate 1–5a、Known Issues C-P-1..7、归因 triage。

本文件只增加 Morph 增量：跨页 shape-name 绑定、Scene Actors vs content 前缀、ghost 纪律、`transition=morph` 自动前缀 quirks、52-style 视觉库查找、Gate 5b morph 扩展。

基座已覆盖时，写 `→ see pptx §X`。未读过 pptx Skill 时先加载再继续。

### Morph 身份（相对 pptx 的 delta）

- **跨页同名绑定。** Morph 引擎用相邻页 **identical**（byte-identical）的 `name=` 配对，并插值 position / size / rotation / fill / opacity。无同名 ⇒ 无动画，静默 fade。这是工作流纪律，不是新 API。
- **命名空间前缀：** `!!scene-*`（持久装饰，极少 ghost）/ `!!actor-*`（会演化并退出）/ `#sN-*`（单页内容，下一页 ghost）。**先规划名字再 `add`。**
- **Ghost 位置 `x=36cm`**（画布右缘外）。禁止删除 `!!` 形状——移出画布才能保留退出动画。
- **`transition=morph` 自动前缀 quirk。** 设置后宿主可能给该页每个 shape 名前加 `!!`（`#s1-title` → `!!#s1-title`）。`@name=` 仍可解析（后缀/前缀容错）；读回名称是加前缀后的形式。见 §Known Issues M-1。
- **相邻页空间多样性。** 位移 ≥ 5cm 或旋转 ≥ 15° 或尺寸 delta ≥ 30%，且至少 3 个配对 shape 发生变化——否则 Morph 插值“看不见”。
- **Renderer 现实。** Morph 在 PowerPoint 365 / Keynote / WPS 中渲染；LibreOffice / 多数 web viewer 常退化为 fade。属 `[RENDERER-BUG]`，不是 Skill 缺陷。

### Reverse handoff

无跨页运动的 deck（board review / sales / all-hands / training）→ 留在 pptx 基座。融资叙事且不要 Morph → `pitch-deck`。仅当用户明确要求 morph / smooth transitions / continuous animation，且 ≥2 连续页共享可变换视觉元素时使用本 Skill。“Animated deck”若仅指单页 entrance animation → pptx §Animations，不是 Morph。


## 宿主工具与执行规范

Agent 通过宿主 Tool JSON 调用 Office 工具，不使用 shell / bash / heredoc。

### Run vs Batch

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- Batch 后必须独立用 Run 做 readback / `view` / screenshot / `validate`。不要把 Batch 成功当成视觉交付。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。
- props 仅允许 string / number / bool。

### 文本语义（宿主侧）

1. Path 含 `[N]` 原样传入 Tool 参数。
2. props 中的 `$` 直接写在 JSON 字符串内（无 shell 展开）。写入后用 Run `view text` 核对货币符号是否保留。
3. 需要真实换行时使用 Help 确认的 paragraph / soft-break 机制，不要手写字面量 `\\n` 当可见换行。


### Morph 特有注意

- `!!` / `#` 在 JSON props 中按字面写入；无 shell history expansion。
- 价格 / 指标文本如 `$9/mo` 直接写在 JSON 字符串中；写入后 Run `view text` 确认没有变成 `/mo` 或孤立 `.`。
- 多 shape 页优先 Batch；读回 / 视觉审计永远用独立 Run。

## 两个原语

- **Scene Actors** = 持久 `!!` 命名 shape，靠同名跨页插值。
- **Choreography** = 谁移动、谁进入、谁退出、在哪一对页上发生——写在 §Morph Pair Planning 表里，**写代码之前完成**。

**Speaker notes 规则。** 每张内容页（非 cover / closing）必须有 notes。缺 notes = 不可交付（继承 pptx H7）。Morph deck 往往视觉极简，notes 承担旁白。

## What is Morph?（核心机制）

PowerPoint Morph 通过相邻页同名 shape 插值属性产生平滑运动：

```
Slide 1: name="!!scene-ring" x=5cm  width=8cm   fill=E94560 opacity=0.3
Slide 2: name="!!scene-ring" x=20cm width=12cm fill=E94560 opacity=0.6
         ↓  slide 2 transition=morph
Result:  ring 平滑移动、放大、加深约 1 秒
```

Morph 只在 slide N+1 带 `transition=morph` 时运行。创建时写入或事后 `set`。省略则静默无动画。

### 三前缀命名（不可协商）

| Prefix | Role | Lifecycle | Example |
|---|---|---|---|
| `!!scene-*` | 背景 / 装饰，整 deck 持续 | 设一次，调位置/尺寸产生运动；**极少 ghost** | `!!scene-ring` |
| `!!actor-*` | 内容 / 前景，跨 section 演化 | 在 N 引入，N+1…修改，退出页 **ghost 到 x=36cm** | `!!actor-metric` |
| `#sN-*` | 单页内容（标题、bullet、caption） | N 新建，N+1 **ghost 到 x=36cm** | `#s1-title` |

**硬规则：** `!!scene-*` 与 `!!actor-*` 名字不得冲突（如 `!!scene-card` + `!!actor-card`）。改成 `!!scene-card-bg` vs `!!actor-card-content`。

**Charts 可参与配对。** 同名 `!!` chart 框会插值位置/尺寸；框内绘制数据不会插值。需要条形“增长”叙事时：接受 chart fade-in，或手工建 `!!actor-bar-K` 矩形并跨页变化 width/height/fill。

**Ghost 累积是静默的。** `!!` shape 一旦出现，后续 morph 页会继续可见，除非显式移到 `x=36cm`。结构 helper / final-check **不会**抓住可见区残留——只有 Gate 5b screenshot / view 审计会。

**同时刻约束。** 同一 morph pair 内所有 `!!` shape **同时**运动。要 A 先于 B：插入中间 keyframe 页。

### Paired / Enter / Exit

| Behavior | Slide A | Slide B | Who carries `!!`? |
|---|---|---|---|
| Paired morph | 有 `!!foo` | 有 `!!foo` | 两边同名 |
| Enter | — | 有 `!!foo` | 仅目标 |
| Exit via ghost | 可见 x | 同名 `x=36cm` | 两边；B 在画布外 |

**退出的是 outgoing 内容。** 忘记 ghost 时名字从 B 消失 = 无动画的 plain fade。永远显式 ghost。

## Morph Pair Planning（编码前必做）

若 audience / purpose / narrative 不清楚，先读 `references/decision-rules.md` 产出 `brief.md`。没有叙事脊柱的 morph 会塌成“带运动的幻灯片”。

在 `brief.md` 中写完整表，再开始 `add`。中途改名是 ghost 累积 bug 的头号原因。

| Pair | Slide A | Slide B | Actors | Ghost on B |
|---|---|---|---|---|
| 1→2 | ring 居中 5cm，`#s1-title` 可见 | ring → x=20cm，8→12cm；`#s2-subtitle` 出现 | `!!scene-ring` | `#s1-title` → 36cm |
| 2→3 | `!!actor-feature-box` 14cm | box 6cm，`!!actor-metric` 进入 | ring + feature + metric | `#s2-subtitle` |
| 3→4 | section A | section B divider | — | 上节 actors + `#s3-*` |

**规划规则：**

1. 预先决定全部 `!!` 名字——两边必须字节级相同。
2. 每个 `!!` 分类为 scene 或 actor；actor 必须有退出页。
3. **Section 边界：** 新 section 首页 ghost 掉上一节全部 `!!actor-*`；只留 `!!scene-*`。
4. 表未完成不要开工。计划变更则重画表并复核受影响页。

## Morph Recipes（4 patterns）

四种模式覆盖约 95% morph deck。每个 recipe 用 Tool JSON；路径示例统一 `/workspace/deck.pptx`。

### 生命周期：create / open

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": ["/workspace/deck.pptx"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "open",
    "command_arguments": ["/workspace/deck.pptx"]
  }
}
```

### (a) Single-element morph — size / position

**视觉结果。** Slide 1 居中 48pt hero；Slide 2 缩到 24pt 并移到左上，让出舞台给新正文。

Slide 1 hero（Batch ≥3 ops）：

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
      "name=!!actor-headline",
      "--prop",
      "text=The one idea",
      "--prop",
      "x=4cm",
      "--prop",
      "y=8cm",
      "--prop",
      "width=26cm",
      "--prop",
      "height=3cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=48",
      "--prop",
      "bold=true",
      "--prop",
      "color=FFFFFF",
      "--prop",
      "align=center",
      "--prop",
      "fill=none"
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
      "notes",
      "--prop",
      "text=Open with the single idea; motion will carry it into evidence."
    ]
  }
}
```

Slide 2 morph + body：

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
      "background=1E2761",
      "--prop",
      "transition=morph"
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
      "name=!!actor-headline",
      "--prop",
      "text=The one idea",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=1cm",
      "--prop",
      "width=12cm",
      "--prop",
      "height=1.5cm",
      "--prop",
      "font=Georgia",
      "--prop",
      "size=24",
      "--prop",
      "bold=true",
      "--prop",
      "color=FFFFFF",
      "--prop",
      "align=left",
      "--prop",
      "fill=none"
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
      "name=#s2-body",
      "--prop",
      "text=Here is the supporting evidence.",
      "--prop",
      "x=1.5cm",
      "--prop",
      "y=5cm",
      "--prop",
      "width=30cm",
      "--prop",
      "height=2cm",
      "--prop",
      "font=Calibri",
      "--prop",
      "size=20",
      "--prop",
      "color=CADCFC",
      "--prop",
      "fill=none"
    ]
  }
}
```

### (b) Multi-element coordinated morph

**视觉结果。** 三个 scene actors 跨页平移/放大，像 camera pan；`#sN-*` 负责文案进出。

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/deck.pptx",
    "operations": [
      {
        "command": "add",
        "parent": "/slide[1]",
        "type": "shape",
        "props": {
          "name": "!!scene-ring",
          "preset": "ellipse",
          "fill": "E94560",
          "opacity": 0.3,
          "x": "5cm",
          "y": "3cm",
          "width": "8cm",
          "height": "8cm"
        }
      },
      {
        "command": "add",
        "parent": "/slide[1]",
        "type": "shape",
        "props": {
          "name": "!!scene-dot",
          "preset": "ellipse",
          "fill": "0F3460",
          "x": "28cm",
          "y": "15cm",
          "width": "1cm",
          "height": "1cm"
        }
      },
      {
        "command": "add",
        "parent": "/slide[1]",
        "type": "shape",
        "props": {
          "name": "#s1-title",
          "text": "Anchor composition",
          "x": "2cm",
          "y": "1cm",
          "width": "28cm",
          "height": "2cm",
          "size": 36,
          "bold": true,
          "color": "FFFFFF",
          "fill": "none"
        }
      }
    ],
    "stop_on_error": true
  }
}
```

Slide 2：确保 `transition=morph`，actors 位移 ≥5cm，并 ghost `#s1-title`：

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/deck.pptx",
    "operations": [
      {
        "command": "set",
        "path": "/slide[2]",
        "props": {
          "transition": "morph"
        }
      },
      {
        "command": "add",
        "parent": "/slide[2]",
        "type": "shape",
        "props": {
          "name": "!!scene-ring",
          "preset": "ellipse",
          "fill": "E94560",
          "opacity": 0.6,
          "x": "20cm",
          "y": "2cm",
          "width": "12cm",
          "height": "12cm"
        }
      },
      {
        "command": "add",
        "parent": "/slide[2]",
        "type": "shape",
        "props": {
          "name": "!!scene-dot",
          "preset": "ellipse",
          "fill": "0F3460",
          "x": "3cm",
          "y": "16cm",
          "width": "1.5cm",
          "height": "1.5cm"
        }
      },
      {
        "command": "add",
        "parent": "/slide[2]",
        "type": "shape",
        "props": {
          "name": "#s1-title",
          "x": "36cm",
          "y": "1cm",
          "width": "28cm",
          "height": "2cm",
          "fill": "none"
        }
      }
    ],
    "stop_on_error": true
  }
}
```

独立 Run 读回两侧 children 名称，确认 `!!scene-ring` / `!!scene-dot` 字节级一致：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[1]",
      "--depth", "1",
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
      "/workspace/deck.pptx",
      "/slide[2]",
      "--depth", "1",
      "--json"
    ]
  }
}
```

### (c) Continuous multi-slide morph（story arc）

**视觉结果。** 5 页连续故事：2 个 scene actors 漂移；每页刷新 `#sN-*` 并在下一页 ghost。对 5+ 页：先在 brief 写全 pair 表，再按页循环：clone-equivalent（新建页 + 重放 scene actors）→ ghost 旧内容 → add 本页内容 → set transition。

中间页 ghost 旧 actor 的 Run（1–2 mutation；更多用 Batch）：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[3]/shape[@name=!!actor-metric]",
      "--prop", "x=36cm"
    ]
  }
}
```

**何时用手写 vs 重复 Batch。** 2–3 页用手写 recipe a/b 更清晰；5+ 页用重复“clone actors + ghost + add content”节奏，并在每页后独立 `validate` / `view`。

### (d) Morph + fade hybrid

Morph 负责 `!!scene-*` 连续运动；新 `#s2-card` 用 Help 确认的 entrance animation（≤600ms，无 bounce/swivel/fly-from-edge）与 morph 同屏出现。

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
      "shape",
      "--prop",
      "name=#s2-card",
      "--prop",
      "preset=roundRect",
      "--prop",
      "fill=F5F7FA",
      "--prop",
      "line=none",
      "--prop",
      "x=2cm",
      "--prop",
      "y=12cm",
      "--prop",
      "width=10cm",
      "--prop",
      "height=5cm"
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
      "/slide[2]/shape[@name=#s2-card]",
      "--prop",
      "animation=fade-entrance-300-with"
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
      "text=Call out the new card while the ring continues to move."
    ]
  }
}
```

读回 animation（注意 trigger 后缀可能被丢掉）：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": [
      "/workspace/deck.pptx",
      "/slide[2]/shape[@name=#s2-card]",
      "--json"
    ]
  }
}
```

## Choreography — 动画类型与错峰

| Animation type | 如何实现（A→B） |
|---|---|
| Simple move | 同名、同尺寸、不同 x/y |
| Scale | 同名、不同 width/height |
| Move + scale | x/y/width/height 同时变 |
| Color / opacity | 同名、不同 fill 或 opacity |
| Rotation | 同名、不同 rotation（最短弧） |
| Font size | 同名、不同 size；跨 viewer 不稳——建议同时改 width/height 或位移 |
| Enter | 仅 B 有 |
| Exit | 仅 A 有；或两边有但 B ghost |

**错峰：** 拆成两对 + 中间 keyframe。不要用 shape `animation=` 伪造 Morph 错峰——Morph 先于 per-shape animation。

**Good-enough variety：** 主导配对 shape 至少改 {x,y,width,height,rotation,fill,opacity} 中的 3 项，且位移 ≥5cm 或旋转 ≥15° 或尺寸 ≥30%。

**Gate 5b-morph-2 更严：** ≥3 个不同的 `!!` shape 各自至少改一项几何/字号。品牌常量（固定 header / footer / logo）**不计入** 3-shape 配额。

**Deck 节奏：**

- 8–10 页：3–5 个 morph 时刻，可成簇。
- 12–18 页：总共 3–5 个，每隔 4–6 页；section divider 用 morph 作章节标点。
- 18+ 页：3 幕结构；幕间 1 个长 section morph + 每幕 2–3 个安静 morph。

## Scene-actor spatial rule

**安全区：**

```
Top-right:   x ≥ 24cm, y ≤ 6cm
Bottom-right: x ≥ 24cm, y ≥ 12cm
Bottom-left:  x ≤ 2cm,  y ≥ 12cm
Ghost:        x ≥ 33.87cm（显式用 36cm）
```

避免把高 opacity actor 停在 content core `x=2~28cm, y=3~16cm`。可穿越，但不要停留除非它就是内容。

放置前先 Run `get /slide[N] depth 1`，确认不与 `#sN-*` bbox 重叠；若重叠则 `opacity≤0.15` 或移到安全区。

## Style library lookup

读 `references/style-catalog.md`（由 52-style INDEX 提炼）。**查逻辑，不抄坐标。**

1. 按 mood / brand hex 选 1 个 family。
2. 读 philosophy（palette / type / signature gesture）。
3. 用 pptx 网格数学落地；只借用 palette 与 signature，不复制 demo 坐标。

## Ghost Discipline & Actor Lifecycle

每张页的 shape 列表独立。在 slide N ghost 不等于 slide N+1 自动 ghost。

1. Slide N：引入 `!!actor-ring`（可见）。
2. Slide N+1：加新内容后，立刻把不应出现的 actors ghost 到 `x=36cm`。
3. 后续每页重复 ghost，直到再次需要可见。

检测：统计 `x≥34cm` 的 ghost。经验上 ≤50（约 4–5 actors × 10–12 页）。显著超标 = M-2 ghost accumulation。用 Gate 5b 循环 + screenshot 定位。

## Delivery Gate

### Gate 1–5a（继承 pptx）

Schema validate、token grep、hyperlink rPr、slide-order、dark-on-dark。全部打印 OK 才继续。Morph deck 与普通 pptx 有相同 token/schema/order 风险。

### Gate 2 morph addendum — 价格 / 指标泄漏

Run `view text`，人工/程序检查是否出现被吃掉的 `$9/mo`、`${VAR}`、字面量 `\n`、孤立 `.`。修复：重新 `set` 正确文本 props。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": ["/workspace/deck.pptx", "text"]
  }
}
```

### Gate 5b — Visual audit（强制）+ morph 扩展

Run `view html` 并 Read HTML；必要时 screenshot。对每页回答 pptx Gate 5b 问题，再加：

- **5b-morph-1：** 应退出的 `!!actor-*` 是否仍 `x < 33.87cm`？
- **5b-morph-2：** 每个 morph pair 是否 ≥3 个不同 `!!` shape 几何变化？
- **5b-morph-3：** 相邻页是否共享 ≥2 个精确 `!!` 名？
- **5b-morph-4：** `#s(N-1)-*` 是否在 slide N 仍可见？

`query` **没有** `^=` 前缀算子。用 `get depth 1` + 结果过滤 `startswith("!!actor-")` / `!!scene-`，不要裸 `startswith("!!")`（会误伤自动前缀内容）。

任一 5b-morph 失败 → REJECT交付。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": ["/workspace/deck.pptx", "html"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": ["/workspace/deck.pptx"]
  }
}
```

## Renderer honesty

**会渲染 Morph：** PowerPoint 365、Keynote、WPS、PowerPoint Online。
**不会（常退化为 fade）：** LibreOffice Impress、Google Slides web、多数 HTML/SVG、`view html`（结构 only）。静态截图无法证明运动——用 Gate 5b 证明配对正确，用目标 viewer 证明运动质量。

## Common Morph Pitfalls

| Pitfall | Correct |
|---|---|
| `!!scene-card` + `!!actor-card` | 跨前缀唯一名 |
| 中途改名 | 停工、重画 pair 表 |
| actor 进 content core 无退出计划 | pair 表先写 exit |
| 忘记每页 re-ghost（M-2） | 每页加内容后立刻 ghost |
| 忘记 `transition=morph` | Gate 5b-morph-2 捕获；`set` 补上 |
| 以为 `@name=` 在 morph 页失效 | 仍可用；读回带 `!!` |
| 相邻页视觉相同 | 至少 3 shape 有意义变化 |
| 用 per-shape timing 错峰 | 拆中间 keyframe |
| 在 LibreOffice 测 Morph | `[RENDERER-BUG]` |
| 删除而非 ghost | 永远 ghost |
| 字面量 `<a:br/>` 塞进 text | 用逐段 paragraph add（M-6） |
| `shape[name^=!!actor-]` | 无效 selector；改 get+过滤 |

## Known Issues（M-1..M-6）

基座 C-P-1..7 全部适用 → pptx §Known Issues。

| # | Symptom | Workaround |
|---|---|---|
| **M-1** | `transition=morph` 后名字自动加 `!!`；裸 `startswith("!!")` 误匹配 | `@name=` 仍解析；过滤用 `!!actor-` / `!!scene-` |
| **M-2** | ghost 累积：actor 在后续页残留 | 每页显式 ghost；Gate 5b；ghost 计数 >50 拒绝 |
| **M-3** | section 边界旧 actor 残留 | 新 section 首页 ghost 全部上节 actors |
| **M-4** | 发明 `morph.duration=` / `transition.delay=` | 用 Help 确认的 `transition=morph-slow` / `-fast` / `morph-<ms>`；超出 shorthand 的 raw XML 调整需宿主 **approval** 的 `raw-set`，且逐步 Run（依赖 relationship/path），**永不**进 Batch JSON |
| **M-5** | LibreOffice / web = fade | 在 PowerPoint 365 / Keynote / WPS 测 |
| **M-6** | text 中字面量 `<a:br/>` | 每行一个 paragraph add |

## Outputs & delivery

交付三件套：

1. `/workspace/<topic>.pptx` — `validate` 干净。
2. 可复现构建说明（宿主 Tool 序列，不是 shell 脚本要求）。
3. `brief.md`：audience / purpose / narrative / style；slide outline；Morph Pair Planning 表。

告知用户：请在 PowerPoint 365 / Keynote / WPS 打开查看 Morph；其他 viewer 可能静态或 fade。

## Adjustments after creation

`swap` / `move` 重排 morph 配对页后，必须重跑 Gate 5b-morph-3。交付前跑完整 Gate 1–5b，并在目标 viewer 看至少一对完整 morph。

## References

- `references/decision-rules.md` — Pyramid / SCQA / brief schema
- `references/design-rules.md` — Scene Actors / page types / morph typography
- `references/style-catalog.md` — 52-style 家族与 use-case 查找
- 基座 `pptx` Skill — 视觉底线、网格、Gate 1–5a、C-P-1..7
