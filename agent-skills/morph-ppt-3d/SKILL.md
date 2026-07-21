# Morph PPT 3D Agent Skill

Author: xiesq, 2026-07-15

## ⚠️ Help 优先规则

**本 Skill 在 morph-ppt 上增加 GLB 3D model、cinematographic camera、模型内容布局与增强视觉系统。** 不确定的 model / camera / shape 字段，先查 Help，禁止发明。

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "pptx",
      "3dmodel"
    ]
  }
}
```

若 Help 不存在 `3dmodel` / model 元素，停止并 **degrade** 为 2D Morph（加载 `morph-ppt`），不要发明 props。

先加载依赖：

```json
{
  "tool": "{{OFFICE_LOAD_SKILL_TOOL}}",
  "arguments": {
    "command_arguments": [
      "morph-ppt"
    ]
  }
}
```

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

`{{OFFICE_LOAD_SKILL_TOOL}}` 只读取指导。

## 心智模型与继承

本 Skill **扩展** `morph-ppt`。命名、ghosting、design、verification、Gate 5b morph **全部适用**。本文件只覆盖 3D 增量 + 更具体的 palette / font / layout guardrails。

### Use when

- 用户要带 `.glb` 3D model 的 Morph `.pptx`。

### Reverse handoff

- 只要 Morph、无 3D → `morph-ppt`。
- 无跨页运动 → `pptx`。
- 融资叙事无 Morph → `pitch-deck`。

## 宿主工具与执行规范

Agent 通过宿主 Tool JSON 调用 Office 工具，不使用 shell / bash / heredoc。

### Run vs Batch

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- Batch 后必须独立用 Run 做 readback / `view` / screenshot / `validate`。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。
- props 仅允许 string / number / bool。


### 3D 特有执行注意

- `.glb` 是 **外部输入**，需要宿主文件 approval / sandbox mount。禁止把 Base64 塞进 props 绕过策略。
- 外部文件路径用真实 path（如 props `path`）；不要伪造内嵌二进制。
- 克隆含 3D model 的页是高风险：冻结 XML / 重名冲突可导致 PowerPoint repair 删除模型。优先每页独立 `add` 同名 model。
- 若必须清理克隆残留，涉及 approval 的 remove / raw 路径用逐步 Run（依赖 relationship/path），**永不**把 `raw-set`/`add-part` 写进 Batch JSON。

## 3D Model Compatibility Gate（生成前）

1. **只支持 `.glb`。** 用户给 `.fbx` / `.obj` / `.blend` / `.usdz` / `.gltf` → 请其先转为 `.glb`（例如 Blender export）。
2. 没有模型 → 走 Model Discovery Flow（下方）。
3. `.glb`、`.pptx`、构建说明应在同一工作目录语义内（宿主可访问的批准路径）。
4. 文件存在、非空、扩展名 `.glb`、建议 < 50MB。

## Model Discovery Flow（用户无模型时）

主动帮助找匹配模型，而不是只贴网站列表。

### Step 1: 理解话题并建议模型方向

| Topic type | Model suggestion | Example |
| --- | --- | --- |
| Product/brand | 真实产品或近似物 | coffee brand → cup / machine / bean |
| Animal/character | 动物或吉祥物 | fox mascot → fox |
| Architecture/space | 建筑 / 室内 | new office → building / interior |
| Vehicle/transport | 交通工具本身 | EV launch → car |
| Food/cooking | 菜肴 / 食材 | Japanese food → sushi / ramen |
| Tech/gadget | 设备 | phone launch → phone / laptop |
| Nature/science | 对象本身 | solar system → planet |
| Abstract concept | 象征物 | teamwork → gears / bridge |

告诉用户：话题是 [X]，建议模型是 [description]，并给出可用来源。

### Step 2–7: 搜索 / 确认 / 下载纪律

- 可代表用户搜索候选（Sketchfab / Poly Pizza / Khronos Sample Assets），但**下载前必须用户确认**。
- 展示 2–3 个候选：名称、来源、预览链接、许可证、为何匹配。
- License 提醒：CC0 / CC BY 可用；CC BY-NC 仅非商用。
- 用户说 "anything / you decide" → 先澄清话题方向（Tech / Animal / Architecture / Food / Other），再搜。
- 用户自助时给出过滤步骤：Downloadable → glTF/.glb → sort by Likes。
- 提醒 web search 有 token 成本，让用户选 agent-search vs self-service。
- Khronos samples（Duck, Fox, Avocado, BrainStem, CesiumMan, DamagedHelmet, FlightHelmet, Lantern, Suzanne, WaterBottle…）适合 demo。

下载后验证：存在、`.glb`、非空、体积合理。Sketchfab 需登录时，请用户分享文件或改用 Khronos sample。

## Visual Design System（3D enrichment）

### Color Palettes（每 deck 选一套）

| Palette | Primary | Secondary | Accent | Body | Muted |
| --- | --- | --- | --- | --- | --- |
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

Rules: 一色主导 60–70%；浅底用 Body / Muted；深底用 Secondary/`FFFFFF`。更多灵感见 `morph-ppt/references/style-catalog.md`——学 approach，不抄坐标。

### Font Pairings

| Header | Body | Best For |
| --- | --- | --- |
| Georgia | Calibri | Formal business, finance |
| Arial Black | Arial | Bold marketing, launches |
| Calibri | Calibri Light | Clean corporate |
| Cambria | Calibri | Traditional professional |
| Trebuchet MS | Calibri | Friendly tech, startups |
| Impact | Arial | Bold headlines, keynotes |
| Palatino | Garamond | Elegant editorial, luxury |
| Consolas | Calibri | Developer / technical |

### Hard Rules

- **H4** Body ≥ 16pt（chart axis ≤12pt、短 sublabel ≤14pt 且 ≤5 词、footnotes 例外）。放不下就减字/拆页/减卡。
- **H6** 深底（brightness <30%）上全部 body/card/chart/icon 用白或近白（brightness >80%）。
- **H7** 每张内容页必须有 speaker notes。

Notes 示例（Run，1 mutation）：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/morph-3d.pptx",
      "/slide[2]",
      "--type", "notes",
      "--prop", "text=Walk the camera around the product while stating the benefit."
    ]
  }
}
```

### Visual Element Checkpoint

每 3 张内容页至少 1 张含非文本视觉：icon circle、colored block、large stat、chart、gradient background、shape composition。纯文本仅允许 quote / code / pure table。

## 3D Model Insertion Rules

### 每页 fresh add — NEVER clone model

`morph_clone_slide` / 克隆含 model 的页会复制冻结 XML，克隆体不能 Morph。每页独立 `add` **相同 `name`**。

**CRITICAL：** 克隆后同一页可能出现两个同名 model3d → PowerPoint repair 可能删除模型内容。若不得不克隆 scene actors：立刻 remove 克隆 model，再 fresh add。remove 若触发复杂 relationship，用逐步 Run + 宿主 approval，不要塞进 Batch JSON 的 raw-set/add-part。

**推荐：** 先建空页 + morph transition + scene actors，再每页 fresh add model。

Slide 1 + Slide 2 models（Batch ≥3，含 transition）：

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/morph-3d.pptx",
    "operations": [
      {
        "command": "add",
        "parent": "/slide[1]",
        "type": "3dmodel",
        "props": {
          "path": "/workspace/assets/model.glb",
          "name": "!!model-hero",
          "x": "16cm",
          "y": "1cm",
          "width": "16cm",
          "height": "16cm",
          "roty": 0
        }
      },
      {
        "command": "add",
        "parent": "/slide[2]",
        "type": "3dmodel",
        "props": {
          "path": "/workspace/assets/model.glb",
          "name": "!!model-hero",
          "x": "0.5cm",
          "y": "1cm",
          "width": "18cm",
          "height": "17cm",
          "roty": 50
        }
      },
      {
        "command": "set",
        "path": "/slide[2]",
        "props": { "transition": "morph" }
      }
    ],
    "stop_on_error": true
  }
}
```

### Controllable properties

| Property | What it does | Notes |
| --- | --- | --- |
| `x`, `y` | Position | Standard slide coordinates |
| `width`, `height` | Frame size | Model renders inside frame |
| `name` | Shape name | Must be identical across slides for Morph |
| `roty` | Y-axis rotation | Primary storytelling axis |
| `rotx` | X-axis tilt | Range about -25 to +40（以 Help 为准） |
| `rotz` | Z-axis roll | Rarely needed |

### Do NOT manually set

- `meterPerModelUnit` — auto from GLB bbox
- `preTrans` — auto centering
- `camera` depth/position — auto fit
- Never use approval-gated `raw-set` on 3D transform parameters unless Help + user explicitly require，且逐步 Run

## Model-Content Layout

### Core Principle

Model IS the protagonist. Text supports the model; model does not decorate the text.

### Size Contrast Rule（强制）

相邻页 model 面积比 ≥ 1.5x 或 ≤ 0.67x。面积 = width × height。禁止 consecutive similar sizes。

| Size tier | Width | Height | Area (approx) | When |
| --- | --- | --- | --- | --- |
| XL (bleed) | 28-36cm | 22-28cm | 600-1000 | Close-up, extends beyond edges |
| L (hero) | 18-24cm | 15-19cm | 270-456 | Title, closing, drama |
| M (split) | 13-17cm | 12-16cm | 156-272 | Standard content + text |
| S (accent) | 5-10cm | 5-10cm | 25-100 | Data-heavy, model as icon |

### Layout Patterns (6)

**A — Model right, content left** Content x=1-14cm; model x=15-20cm, w 14-18cm.

**B — Model left, content right** Model x=0-2cm; content x=18-32cm.

**C — Model centered, text overlay** Model 18-24cm centered; text top or bottom.

**D — Model small corner** Model 5-10cm; content fills rest.

**E — Model as backdrop** XL 28-36cm, partially cropped; high-contrast text overlay.

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/morph-3d.pptx",
    "operations": [
      {
        "command": "add",
        "parent": "/slide[3]",
        "type": "3dmodel",
        "props": {
          "path": "/workspace/assets/model.glb",
          "name": "!!model-hero",
          "x": "-2cm",
          "y": "-2cm",
          "width": "38cm",
          "height": "24cm",
          "roty": 45,
          "rotx": 10
        }
      },
      {
        "command": "add",
        "parent": "/slide[3]",
        "type": "shape",
        "props": {
          "name": "#s3-quote",
          "text": "Key insight here",
          "x": "3cm",
          "y": "7cm",
          "width": "28cm",
          "height": "5cm",
          "size": 44,
          "bold": true,
          "color": "FFFFFF",
          "fill": "none"
        }
      },
      {
        "command": "set",
        "path": "/slide[3]",
        "props": { "transition": "morph" }
      }
    ],
    "stop_on_error": true
  }
}
```

**F — Model bleed edge** Partial off-screen implies more beyond the frame.

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/morph-3d.pptx",
    "operations": [
      {
        "command": "add",
        "parent": "/slide[4]",
        "type": "3dmodel",
        "props": {
          "path": "/workspace/assets/model.glb",
          "name": "!!model-hero",
          "x": "20cm",
          "y": "-1cm",
          "width": "24cm",
          "height": "22cm",
          "roty": 70
        }
      },
      {
        "command": "add",
        "parent": "/slide[4]",
        "type": "shape",
        "props": {
          "name": "#s4-title",
          "text": "There is more beyond the frame",
          "x": "1.5cm",
          "y": "2cm",
          "width": "16cm",
          "height": "3cm",
          "size": 32,
          "bold": true,
          "color": "FFFFFF",
          "fill": "none"
        }
      },
      {
        "command": "add",
        "parent": "/slide[4]",
        "type": "notes",
        "props": { "text": "Tease the next section with a bleed framing." }
      }
    ],
    "stop_on_error": true
  }
}
```

### Layout Progression 示例

```
S1 C L → S2 E XL (≥1.5x) → S3 A M → S4 F L → S5 D S → S6 B M → S7 C L
```

Never repeat the same pattern on consecutive slides.

### Text Layout Safety（强制）

1. Title wrap → body_y = title_y + title_height + 0.5cm。
2. 用慷慨高度：title 3–4cm，body 6–8cm，bullets 8–10cm。
3. Model frame 与 text gap ≥ 1cm。
4. Pattern C：text 在 top (y=0.5–2) 或 bottom (y=14–17)，不在 model 垂直中部 (y=3–13)。
5. 每页后 Run `get depth 1` 查重叠。

### Model Bleed Guidelines

Bleed 适合：对称物、大平面、裁非关键部位。
不适合：角色/动物裁耳尾肢、小细节模型、裁掉最可识别特征。
角色/动物：保持完整可见，用 L→M→S 与 `rotx` 做节奏。

## Camera Language

三工具：**roty**（orbit）、**rotx**（tilt）、**width/height**（zoom）。

### Shot Types（每 deck ≥3 种）

| Shot | Size | rotx | When |
| --- | --- | --- | --- |
| Establishing | L 18-24 | 0-5 | Title, intro, closing |
| Three-quarter beauty | L 16-20 | 5-10 | Hero first impression |
| Close-up | XL 28-36 cropped | 0-10 | Feature detail |
| Bird's eye | M 13-17 | 25-40 | Structure overview |
| Low angle | L 16-20 | -15 to -25 | Power, drama |
| Side profile | M 13-16 | 0 | Form factor |
| Over-the-shoulder | S 5-10 | 10-15 | Data-heavy accent |

### Content-Driven Camera

- Front design → Close-up, roty=0, XL
- Side profile → roty=90, M
- Internal structure → Bird's eye, roty=30, rotx=35, M
- Power → Low angle, roty=20, rotx=-20, L
- Data & specs → Over-the-shoulder, roty=60, S corner

### Rotation Rules

1. Adjacent roty delta 30–90°（<30=jitter, >90=disorienting）
2. Overall roty direction consistent（no back-and-forth）
3. rotx range ~-25..+40；adjacent rotx delta ≤20
4. Total arc 180–360° across deck

### Example Shot Plan

| Slide | Shot | roty | rotx | Size | Pattern |
| --- | --- | --- | --- | --- | --- |
| 1 | Three-quarter | 30 | 8 | L 20×17 | C |
| 2 | Close-up | 0 | 5 | XL 30×24 | E |
| 3 | Side | 80 | 0 | M 15×14 | A |
| 4 | Bird's eye | 120 | 35 | M 14×13 | B |
| 5 | Low angle | 170 | -20 | L 20×18 | F |
| 6 | OTS | 220 | 10 | S 8×7 | D |
| 7 | Establishing | 320 | 5 | L 20×17 | C |

## Workflow Integration with morph-ppt

### Phase 2 — Planning

在 `brief.md` 增加 Model Choreography Table：

| Slide | Pattern | Size Tier | Model x,y,w,h | roty | rotx |
| --- | --- | --- | --- | --- | --- |
| 1 | C | L | 7,0.5,20,17 | 30 | 8 |
| 2 | E | XL | -2,-2,38,24 | 0 | 5 |

Verify area ratio before build.

### Phase 3 — Build

1. Create all slides（background + morph transition）
2. Add scene actors on slide 1；为连续性在后续页重放同名 scene（不要靠克隆 model）
3. Fresh add 3D model on EACH slide（同名，不同 roty/position/size）
4. Add content；ghost previous `#sN-*` / exiting actors

Create/open：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": ["/workspace/morph-3d.pptx"]
  }
}
```

### Phase 4 — Verification

在标准 morph 验证之外额外检查：

- 每页恰好一个 `model3d`
- 所有 model 共享相同 `name`
- 相邻面积比 ≥1.5x 或 ≤0.67x
- 无连续相同 layout pattern
- 独立 Run `view` / screenshot / `validate`

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "get",
    "command_arguments": [
      "/workspace/morph-3d.pptx",
      "/slide[2]",
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
    "command_name": "view",
    "command_arguments": ["/workspace/morph-3d.pptx", "html"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": ["/workspace/morph-3d.pptx"]
  }
}
```

## File Placement / Deliverables

恰好 4 类交付：批准的 `.glb`、输出 `.pptx`、可复现构建说明（Tool 序列）、`brief.md`。不要额外 outline.md / quality-report.md 等。规划进 brief，验证输出进 stdout/对话。

## Common traps

| Trap | Fix |
| --- | --- |
| 非 GLB 直接塞 | 先转换 |
| 克隆页保留冻结 model | remove 后 fresh add；或根本不克隆 |
| 相邻尺寸几乎一样 | 强制 1.5x / 0.67x |
| 角色模型 bleed 裁肢 | 改用 size/rotx 节奏 |
| text 压在 model 上 | gap≥1cm 或移到 top/bottom |
| 发明 camera raw 字段 | Help + approval + 逐步 Run |
| Batch 成功当视觉交付 | 独立 screenshot / view / validate |

## Renderer honesty

3D + Morph 的 runtime 观感依赖目标 viewer。`view html` / screenshot 用于结构与布局审计，不能替代 PowerPoint 365 / 支持 3D+Morph 的宿主里的真实预览。明确告知用户。
