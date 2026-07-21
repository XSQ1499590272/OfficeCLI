# Word-Form Agent Skill

Author: xiesq, 2026-07-15

**This skill is INDEPENDENT, not a scene layer on docx.** A form's payload — `<w:sdt>` controls, `<w:ffData>` legacy fields, `<w:fldChar>` mail-merge, `documentProtection` — is a distinct element class from docx's paragraph/heading/style primitives. Its QA is different too: docx's Delivery Gate cares about visual layout and live PAGE fields, this skill's cares about data plumbing (protection enforced / alias+tag / items injected / name ≤ 20 / no underscore anti-pattern). **Reverse handoff:** if the user's document has no fillable fields (report, letter, memo, thesis, proposal), route to `word` or a docx scene skill — don't use this one.

## Help-First Rule

This skill teaches what a real form needs, not every flag. When a prop / alias / enum is uncertain, consult Help BEFORE guessing:

```json
{
  "tool": "{{OFFICE_HELP_TOOL}}",
  "arguments": {
    "command_arguments": [
      "docx",
      "sdt"
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
      "formfield"
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

Help is pinned to the installed CLI version and is authoritative — when this skill and help disagree, **help wins** (the prop set on `sdt` in particular has grown over time; trust Help `docx sdt`, not a hardcoded list).

`{{OFFICE_LOAD_SKILL_TOOL}}` 只读取指导，不安装或修改本地 skills。

## Mental Model & Inheritance

A Word form is a `.docx` plus four OpenXML payload layers plain-docx skills do not touch: **`<w:sdt>`** content controls (types: text / richtext / dropdown / combobox / date / picture / group), **`<w:ffData>`** legacy FormField (still the only way to get a real checkbox — SDT `type=checkbox` is not implemented), **`<w:fldChar>`** complex fields (MERGEFIELD, REF, PAGEREF, SEQ, IF — template-time, not user-fill), and **`documentProtection`** (the lock that makes non-field text read-only in Word — and, on the host tools, `protection=forms` locks non-field content edits (those need `force` or `raw-set`) but still allows form-field (SDT) edits, which is the point of forms protection).

**No inheritance from docx v2.** docx's Delivery Gate (cover-fill %, live-PAGE check) does NOT apply — form QA is `view forms` + `query sdt alias+tag` + `protectionEnforced`.

**Reverse handoff to docx.** Route back to `word` for reports / letters / memos / thesis / pitch decks / any document with no editable fields. Use **this** skill when the document's purpose is data capture or template merge.

## 宿主工具与执行规范

（对应原 CLI「Shell & Execution Discipline」；shell 引号改为 Tool JSON 传参。）

**增量执行。** One mutation at a time when results depend; read output before the next. Every `add` / `set` / `remove` immediately mutates the file. 示例 path 使用 `/workspace/form.docx`。

**Path 与特殊字符：**

- path 原样传入 `command_arguments` / Batch `path`（含 `[N]`），无需 shell 引号。
- `$` 写在 JSON 字符串内即可（无 shell expansion）。
- `after find:<text>`：锚点字符串原样传入 prop；不要把引号写进搜索串本身。

**阶段 0 Batch/Run 契约（固定补丁，非业务加戏）。**

- **Inspect first** with `{{OFFICE_RUN_TOOL}}`（`get` / `query` / `view` / `validate`），再决定是否编辑。
- 同一文件通常 ≥3 个参数已知、相互独立的 mutation 才优先 `{{OFFICE_BATCH_TOOL}}`（`set/add/import/remove/move/swap`；`raw-set/add-part` 仅 prose + 宿主 **approval**，依赖 **relationship**/path 时逐步 Run，**永不**进 Batch JSON）。数量是 guidance heuristic，不是 schema 硬限制。
- 单步、1–2 项、结果依赖、丰富诊断、以及 `create/open/save/close/get/query/view/raw/validate` 用 `{{OFFICE_RUN_TOOL}}`；禁止 `command_name=batch`。
- Batch **不是事务**：`stop_on_error` 只停止后续、**不回滚**。仅在接受 **partial success** 或已有 **discardable** copy 时使用；全有或全无且无副本时不要 Batch 原件。阶段 0 **不会自动创建草稿**、发起审批或原子覆盖。
- Batch JSON 超过 **8192** bytes 只返回 `outputFile` slim envelope；阶段 0 不归一化——避免 read-heavy / 超大 Batch。
- `add/move` 的 `index/after/before` 最多一个；禁止 `add.from` 与 `props` 同时出现。

**`WARNING: UNSUPPORTED` (exit 2) is a silently-wrong element.** The CLI created the element *without* the rejected prop. Any UNSUPPORTED in your build log means a prop name the current CLI does not accept on that element — stop, check Help `docx <element>` for the right prop name (most SDT props such as `items`/`format`/`lock` ARE accepted now; `maxlength` is not), fix the command, and re-run. Do not ship on top.

**`protection=forms` is the LAST structural command.** Once it is set, `protection=forms` locks non-field content (those edits need `force` or `raw-set`) but still allows form-field (SDT) edits — which is the point of forms protection. A non-field content edit (e.g. `set /body/p[N] prop text=`) is refused (`ERROR: Document is protected … use force`); a form-field edit (e.g. `set /body/sdt[N] prop text=` / `prop alias=`) succeeds. Use `Query("editable")` to find fillable fields. So finish all non-field (static layout) edits first, then lock; if you must edit static content afterward, pass `force` (or temporarily clear protection, edit, re-lock).

### `after find:` micro-playbook

`after find:<text>` matches the **first** occurrence. Bad anchor = wrong insertion location, expensive to debug. Three rules:

1. **Anchor must be globally unique.** In bilingual contracts "甲方签字" matches both parties — use a unique phrase like "甲方签字（Service Provider）" or full English title.
2. **After insert, `/body/p[last()]` is unreliable** — the find insertion changes `<w:body>` child order. To continue operating on the new paragraph, read its real paraId: Run `query paragraph json`，取 `.data.results[-1].format.paraId`（不要用 shell `jq`）。
3. **Chinese + full-width parens `（）`** match literally in `find`, but when unsure, Run `view text` first to confirm the exact bytes in the file.

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/form.docx",
      "/body",
      "--type",
      "sdt",
      "--after",
      "find:签字"
    ]
  }
}
```

Trap: first-match hits 甲方 only, 乙方 missed. Fix: two signatories, two unique anchors:

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/form.docx",
    "operations": [
      {"command":"add","parent":"/body","type":"sdt","props":{"alias":"Party_A_Name","tag":"party_a"},"after":"find:甲方签字（Service Provider）"},
      {"command":"add","parent":"/body/p[@paraId=PARA_ID_FROM_QUERY]","type":"sdt","props":{"alias":"Party_A_Title","tag":"party_a_title"}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"乙方签字（Client）","bold":true}}
    ],
    "stop_on_error": true
  }
}
```

（第二步的 `paraId` 须先独立 Run `query paragraph` 取得后再填入；结果依赖时不要整段盲 Batch。）

Inline SDT via `after find:` is added as a child of the matched paragraph, not as a new paragraph — use this when label + SDT must share a line.

## What makes a real form (identity)

A real fillable form requires **structured fields** + **document protection**.

| Approach | Word user sees | CLI-readable | Real form? |
|---|---|---|---|
| SDT controls + `protection=forms` | Gray-bordered fields; rest locked | `query sdt` / `view forms` | **YES** |
| FormField checkbox + `protection=forms` | Real clickable checkbox; rest locked | `query formfield` / `view forms` | **YES** (checkbox only) |
| MERGEFIELD placeholders | `«CustomerName»` merged by downstream engine | `query field` | **YES** (template-time) |
| Underscores `___` / blank lines | Visual-only; whole doc editable | No — no structured fields | **NO** |

**Do not simulate fields with underscores.** `姓名：_______________` produces zero structured data and leaks past every verification. Always use `type=sdt` or `type=formfield`.

**Checkbox is formfield, NOT SDT.** `type=sdt` + `type=checkbox` exits 1 (`SDT type 'checkbox' is not implemented`). Every checkbox in every recipe uses `type=formfield` + `type=checkbox`.

**MERGEFIELD is a separate track.** `view forms` lists SDT + formfield only; `query field` lists complex fields only. Two disjoint inventories; both valid in one file.

## Requirements for Outputs (hard floor)

Every form must satisfy these — Delivery Gate enforces each as an executable check.

1. `protection=forms` enforced (`get /` → `protectionEnforced=True`).
2. Every SDT has both `alias` + `tag`.
3. Every dropdown/combobox has non-empty `items=...` in `view forms`.
4. Every date SDT shows the intended `format=...`.
5. Every locked SDT shows `lock=sdtLocked` / `contentLocked` / `sdtContentLocked` as intended.
6. Zero `WARNING: UNSUPPORTED` in build log.
7. Zero `type=checkbox` on any SDT.
8. Every formfield `name` ≤ 20 characters.
9. Zero underscore-line / blank-line placeholders.
10. Field types match user intent (short text / paragraph / fixed list / list+custom / date / boolean).

## Three Paths (core decision)

SDT props are **first-class on `add`** — confirm the current set with Help `docx sdt`. Today that includes `type, tag, alias, text, items, format, lock, placeholder/placeholderText, date.*` and the SDT types `text / richtext / dropdown / combobox / date / group / picture`. So most forms are pure `add` — no raw-set, no Word template. Two paths remain only for the genuinely-unreachable cases. **Pick the path before writing a single command.**

### Path A — Pure CLI (the default for almost everything)

**Use when**: any text / richtext / dropdown / combobox / date / picture / group SDT — including its options, date format, and lock. Pass the props straight on `add`; verify each persists with `get '/body/sdt[N]' json`.

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/form.docx",
    "operations": [
      {"command":"add","parent":"/body","type":"sdt","props":{"type":"dropdown","alias":"Department","tag":"dept","items":"Engineering,Finance,HR"}},
      {"command":"add","parent":"/body","type":"sdt","props":{"type":"date","alias":"Start Date","tag":"start","format":"yyyy年MM月dd日"}},
      {"command":"add","parent":"/body","type":"sdt","props":{"type":"text","tag":"full_name","alias":"Full Name","text":"Enter full name","lock":"sdtLocked"}}
    ],
    "stop_on_error": true
  }
}
```

# protection comes last, once all fields exist:
# Run set / prop protection=forms

### Path B — CLI + `raw-set` bridge (only an attribute help does NOT expose)

**Use when**: you need an SDT attribute that Help `docx sdt` does not list (e.g. a `<w:listItem>` whose `w:value` must differ from its display text, or a sdtPr child with no prop). `raw-set` is 本 Skill 的 universal OpenXML fallback. For ordinary dropdowns use `items=` (Path A); raw-set here is the exception, not the rule. **仅逐步 Run + 宿主 approval；永不进 Batch JSON。**

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/form.docx",
      "/body",
      "--type",
      "sdt",
      "--prop",
      "type=dropdown",
      "--prop",
      "alias=Department",
      "--prop",
      "tag=dept"
    ]
  }
}
```

Then (approval required):

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "raw-set",
    "command_arguments": [
      "/workspace/form.docx",
      "/document",
      "--xpath",
      "//w:sdt[w:sdtPr/w:tag/@w:val='dept']/w:sdtPr/w:dropDownList",
      "--action",
      "append",
      "--xml",
      "<w:listItem xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:displayText=\"Engineering\" w:value=\"ENG\"/><w:listItem xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:displayText=\"Finance\" w:value=\"FIN\"/>"
    ]
  }
}
```

### Path C — Word template (only what no API reaches)

**Use when**: a **real SDT checkbox** (`type=checkbox` still exits 1 — use a legacy FormField instead, see §Legacy FormField), a `placeholderDocPart` prompt-text part, or custom richtext appearance / cross-part nesting beyond prop reach. Picture and grouped SDTs are NO LONGER here — they add fine via Path A.

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "open",
    "command_arguments": [
      "/workspace/form.docx"
    ]
  }
}
```

（先由人工将 Word 模板复制到目标 path，再 open。）

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": [
      "/workspace/form.docx",
      "forms"
    ]
  }
}
```

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/form.docx",
      "/body/sdt[@sdtId=3]",
      "--prop",
      "text=Jane Smith"
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
      "/workspace/form.docx",
      "/",
      "--prop",
      "protection=forms"
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
      "/workspace/form.docx",
      "/body/sdt[1]",
      "--prop",
      "lock=sdtlocked"
    ]
  }
}
```

### Decision table

| Need | Path | Note |
|---|---|---|
| text / richtext SDT with default string | **A** | `type/alias/tag/text` |
| text SDT that must be locked | **A** | `lock=sdtLocked` works on `add` (and `set`) |
| dropdown / combobox **with options** | **A** | `items="A,B,C"` |
| date SDT with non-default format | **A** | `format="yyyy年MM月dd日"` |
| signature picture SDT, grouped SDT | **A** | `type=picture` / `type=group` add directly |
| dropdown whose stored value ≠ display text | **B** | raw-set append `<w:listItem w:value=…>`（Run + approval） |
| real checkbox | **FormField** | `type=formfield` + `type=checkbox` |
| mail-merge placeholder | **MERGEFIELD** | `type=field` + `fieldType=mergefield` |
| real SDT checkbox / placeholder part / custom appearance | **C** | build skeleton in Word, fill via host tools |

## Quick Start — Path A + FormField (minimal intake form)

Two SDT text fields, one checkbox, protection. Paste and adapt; this is the smallest form worth shipping.

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": [
      "/workspace/intake.docx"
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
      "/workspace/intake.docx"
    ]
  }
}
```

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/intake.docx",
      "/",
      "--prop",
      "title=Employee Onboarding Intake",
      "--prop",
      "docDefaults.font=Calibri",
      "--prop",
      "docDefaults.fontSize=12pt"
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
      "/workspace/intake.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Employee Onboarding Intake",
      "--prop",
      "style=Heading1",
      "--prop",
      "size=20",
      "--prop",
      "bold=true",
      "--prop",
      "spaceAfter=18pt"
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
      "/workspace/intake.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Full Name:",
      "--prop",
      "size=11",
      "--prop",
      "bold=true",
      "--prop",
      "spaceAfter=4pt"
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
      "/workspace/intake.docx",
      "/body",
      "--type",
      "sdt",
      "--prop",
      "type=text",
      "--prop",
      "alias=Full Name",
      "--prop",
      "tag=full_name",
      "--prop",
      "text=Enter full name"
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
      "/workspace/intake.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Start Date:",
      "--prop",
      "size=11",
      "--prop",
      "bold=true",
      "--prop",
      "spaceAfter=4pt"
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
      "/workspace/intake.docx",
      "/body",
      "--type",
      "sdt",
      "--prop",
      "type=date",
      "--prop",
      "alias=Start Date",
      "--prop",
      "tag=start_date"
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
      "/workspace/intake.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Read and agree to employee handbook",
      "--prop",
      "size=11",
      "--prop",
      "spaceAfter=4pt"
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
      "/workspace/intake.docx",
      "/body",
      "--type",
      "formfield",
      "--prop",
      "type=checkbox",
      "--prop",
      "name=agree_handbook",
      "--prop",
      "checked=false"
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
      "/workspace/intake.docx",
      "/body/sdt[1]",
      "--prop",
      "lock=sdtlocked"
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
      "/workspace/intake.docx",
      "/body/sdt[2]",
      "--prop",
      "lock=sdtlocked"
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
      "/workspace/intake.docx",
      "/",
      "--prop",
      "protection=forms"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "close",
    "command_arguments": [
      "/workspace/intake.docx"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "view",
    "command_arguments": [
      "/workspace/intake.docx",
      "forms"
    ]
  }
}
```

（冷启动 / 升级后若有 stale resident，先独立 Run `close`，再 `create`。）

## Path B — raw-set recipes

Three recipes cover almost every complex-attr need on SDT forms. **全部 raw-set 仅 Run + approval；永不进 Batch JSON。**

### B1 — Dropdown items (append)

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/form.docx",
      "/body",
      "--type",
      "sdt",
      "--prop",
      "type=dropdown",
      "--prop",
      "alias=Department",
      "--prop",
      "tag=dept"
    ]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "raw-set",
    "command_arguments": [
      "/workspace/form.docx",
      "/document",
      "--xpath",
      "//w:sdt[w:sdtPr/w:tag/@w:val='dept']/w:sdtPr/w:dropDownList",
      "--action",
      "append",
      "--xml",
      "<w:listItem xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:displayText=\"Engineering\" w:value=\"Engineering\"/><w:listItem xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:displayText=\"Finance\" w:value=\"Finance\"/><w:listItem xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:displayText=\"HR\" w:value=\"HR\"/>"
    ]
  }
}
```

Verify: Run `get '/body/sdt[1]'` — expect `type=dropdown items=Engineering,Finance,HR`.

**Template.** Swap `<TAG>` / `<LABEL>` / `<VALUE>` only. `xmlns:w=...` is required on every root `<w:listItem>` — raw-set does not inherit namespace prefixes. Chain multiple `<w:listItem>`s in one call; option order is preserved.

### B2 — Combobox items (same as B1, different xpath tail)

Same as B1 but xpath tail is `w:comboBox` vs `w:dropDownList`. Combobox lets the user type custom input; dropdown does not.

### B3 — Date format (setattr)

Prefer Path A `format=` on `add`. If you must setattr:

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "raw-set",
    "command_arguments": [
      "/workspace/form.docx",
      "/document",
      "--xpath",
      "//w:sdt[w:sdtPr/w:tag/@w:val='contract_start']/w:sdtPr/w:date/w:dateFormat",
      "--action",
      "setattr",
      "--xml",
      "w:val=yyyy年MM月dd日"
    ]
  }
}
```

`setattr` replaces one attribute — do not quote the value inside `xml`. Only `w:val` is touched; the `<w:dateFormat>` wrapper is preserved.

### raw-set actions & errors

| `action` | Form use |
|---|---|
| `append` | Insert new child at end of target (B1, B2 — listItem) |
| `setattr` | Change one attribute; `xml "key=value"` (B3 — dateFormat/@val) |
| `replace` | Replace entire target (rare — reset a full `<w:date>` wrapper) |
| `remove` | Delete the target (clear options before re-populate) |

| Symptom | Fix |
|---|---|
| `raw-set: 0 element(s) affected` | XPath did not match. Check the `tag` value and whether the SDT is block or inline. Fall back to Run `raw /document` to read the real XML. |
| `Error: prefix 'w' is not defined` | Missing `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` on the fragment — every root element in `xml` needs it. |
| Items readback empty after append | `<w:dropDownList/>` must already exist (Path A `type=dropdown` ensures this). If absent, append has nowhere to insert. |
| `VALIDATION: N new error(s) introduced` on same line as success | Your append introduced a schema-invalid child. Treat as stop-and-fix even though `raw-set` exits 0. |

## Path C — Word template workflow

For fields CLI cannot express (real SDT checkbox, `placeholderDocPart` prompt text, custom richtext styling), build the skeleton once in Word, then fill via host tools.

**One-time in Word:** File → Options → Customise Ribbon → Developer. Developer tab → Insert Picture / Check Box / Grouping Content Control → right-click → Properties → set Title (`alias`) + Tag. Save as `template.docx`.

**Fill via host tools:** open template copy → `view forms` → `set` field values → `protection=forms` → `close`（见上方 Path C JSON）。

## MERGEFIELD (data-driven track)

Help `docx field` declares a `fieldType` enum of ~30 values including `mergefield`, `ref`, `pageref`, `seq`, `if` — all expressible with their typed props. MERGEFIELD coexists with SDT in the same file but is reported by `query field` only; `view forms` does NOT list MERGEFIELDs (they are not user-fillable).

**Canonical MERGEFIELD:**

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/form.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Dear "
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
      "/workspace/form.docx",
      "/body/p[1]",
      "--type",
      "field",
      "--prop",
      "fieldType=mergefield",
      "--prop",
      "name=CustomerName"
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
      "/workspace/form.docx",
      "/body/p[1]",
      "--type",
      "run",
      "--prop",
      "text=, "
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
      "/workspace/form.docx",
      "/body/p[1]",
      "--type",
      "field",
      "--prop",
      "fieldType=mergefield",
      "--prop",
      "name=CompanyName"
    ]
  }
}
```

# Readback: "Dear «CustomerName», «CompanyName»"

**Element-type shortcut** (equivalent): `add … type mergefield prop name=CustomerName`.

### Common field patterns

| Pattern | Call shape |
|---|---|
| Mail-merge placeholder | `type=field` + `fieldType=mergefield` + `name=<FieldName>` |
| Mail-merge with numeric picture (money, percent) | `fieldType=mergefield` + `name=Amount` + `instr='MERGEFIELD Amount \# "#,##0.00"'`. The typed `format` prop is ignored for mergefield (prints a warning) — use `instr` (alias `instruction`). Verify: Run `query field json`；instruction 须含 `\#` 与 picture。 |
| Mail-merge with date picture | `instr='MERGEFIELD StartDate \@ "yyyy-MM-dd"'` |
| Cross-reference to bookmark text | `fieldType=ref` + `name=<BookmarkName>` |
| Cross-reference to bookmark's page number | `fieldType=pageref` + `name=<BookmarkName>` |
| Auto-numbering (Figure 1 / 2 / 3) | `fieldType=seq` + `identifier=Figure` |
| Page number in footer | `fieldType=page` |
| "Page X of Y" | two fields: `fieldType=page` + `fieldType=numpages` |
| Conditional text | `fieldType=if` + `expression` + `trueText` / `falseText` |

### IF conditional (CLI-expressible)

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/form.docx",
      "/body",
      "--type",
      "paragraph",
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
      "/workspace/form.docx",
      "/body/p[last()]",
      "--type",
      "field",
      "--prop",
      "fieldType=if",
      "--prop",
      "expression={ MERGEFIELD Gender } = \"Male\"",
      "--prop",
      "trueText=Mr.",
      "--prop",
      "falseText=Ms."
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
      "/workspace/form.docx",
      "/body/p[last()]",
      "--type",
      "run",
      "--prop",
      "text= "
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
      "/workspace/form.docx",
      "/body/p[last()]",
      "--type",
      "field",
      "--prop",
      "fieldType=mergefield",
      "--prop",
      "name=LastName"
    ]
  }
}
```

Nested wrappers like `{ IF { MERGEFIELD X } = "Y" { REF bm } "fallback" }` are not expressible via prop chaining — drop to raw-set a hand-crafted `<w:fldChar>` / `<w:instrText>` fragment（Run + approval）, or build once in a Word template (Path C).

**Readback.** `query field` lists `/field[N]` + instruction + `fieldType`. `view forms` does NOT list MERGEFIELDs. `get '/body/p[1]'` renders the guillemet-wrapped field name.

## Legacy FormField

Use FormField **when you need a real checkbox**. For text/dropdown, prefer SDT.

Help `docx formfield`: `type` (text/checkbox/check/dropdown), `name` (required, **≤ 20 chars** — OpenXML schema MaxLength; add passes longer but `validate` rejects), `text` (text only, alias `value`), `checked` (checkbox only).

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/form.docx",
      "/body",
      "--type",
      "formfield",
      "--prop",
      "type=checkbox",
      "--prop",
      "name=agree_terms",
      "--prop",
      "checked=false"
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
      "/workspace/form.docx",
      "/body",
      "--type",
      "formfield",
      "--prop",
      "type=text",
      "--prop",
      "name=emp_name",
      "--prop",
      "text=Enter name"
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
      "/workspace/form.docx",
      "/body",
      "--type",
      "formfield",
      "--prop",
      "type=dropdown",
      "--prop",
      "name=dept_select"
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
      "/workspace/form.docx",
      "/formfield[agree_terms]",
      "--prop",
      "checked=true"
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
      "/workspace/form.docx",
      "/formfield[emp_name]",
      "--prop",
      "text=Jane Smith"
    ]
  }
}
```

DROPDOWN formfield — items NOT settable via CLI; use Word template or SDT Path B / Path A `items=`.

FormField paths (`/formfield[N]` or `/formfield[<name>]`) are separate from SDT paths (`/body/sdt[N]`). Both coexist; `protection=forms` covers both.

**Scale.** Tested with 50+ checkboxes in a single document — no practical cap on formfield count; build and `validate` remain clean. `name` ≤ 20 chars (K13) is the only hard constraint.

**Renderer note — formfield checkbox `[RENDERER-BUG]`.** LibreOffice's PDF export occasionally renders the formfield checkbox as `☐☐` (doubled box). Word and WPS render a single clickable box (toggles ☑). This is a LibreOffice renderer quirk, **not a skill or document quality issue** — see K19. Do not attempt workarounds in the form; if an evaluator screenshots a LibreOffice-generated PDF and sees `☐☐`, attribute to `[RENDERER-BUG]`.

## Document protection & lock

### Enabling form protection

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/form.docx",
      "/",
      "--prop",
      "protection=forms"
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
      "/workspace/form.docx",
      "/"
    ]
  }
}
```

# look for: protectionEnforced=True

### Protection modes

| Mode | Word user can | CLI behavior |
|---|---|---|
| `forms` | Fill SDT + formfield only | Form-field (SDT) edits work; non-field content edits need `force` or `raw-set` |
| `readOnly` | Read only | Non-field edits need `force`; raw-set bypasses |
| `comments` | Add comments only | Non-field edits need `force`; raw-set bypasses |
| `trackedChanges` | Edit with tracked changes only | Non-field edits need `force`; raw-set bypasses |
| `none` | Full editing | All ops work |

**KEY:** Document protection restricts Word users AND the host tools. Under `protection=forms`, form-field (SDT) edits succeed; only NON-field content edits are refused with `ERROR: Document is protected (mode: forms). … use force to override`. Use `Query("editable")` to find the fields a Word user could still fill.

### Lock values (settable on `add` and `set`)

下面三种值是互斥替代方案，只选择一个独立 Run，不得连续覆盖：

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": ["/workspace/form.docx", "/body/sdt[1]", "--prop", "lock=sdtlocked"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": ["/workspace/form.docx", "/body/sdt[1]", "--prop", "lock=contentlocked"]
  }
}
```

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": ["/workspace/form.docx", "/body/sdt[1]", "--prop", "lock=sdtcontentlocked"]
  }
}
```

# Omit lock entirely → unlocked (default)

`lock=...` works on `add` as well as `set`. Readback normalises to camelCase (`sdtLocked`) regardless of input case — both accepted.

### lock × `protection=forms` interaction

| lock value | `protection=forms` active | Word user can edit? | Word user can delete control? |
|---|---|---|---|
| (none) | yes | **Yes** | **Yes** |
| `sdtlocked` | yes | Yes | No |
| `contentlocked` | yes | No | Yes |
| `sdtcontentlocked` | yes | No | No |
| block-level SDT wrap `contentlocked` | any | No (wrapped paragraph read-only regardless of protection) | No |
| any | `readOnly` mode | No | No |

### Block-level lock (paragraph-wrapping SDT)

`protection=forms` is document-level — once an admin unprotects, every static paragraph (disclaimer, legal attestation, contract clause) becomes editable again. Master templates need defense-in-depth: wrap the critical paragraph in a block-level `<w:sdt>` with `lock=contentLocked`, so the content stays read-only even after protection is stripped.

做法概要（与原 CLI 相同，工具改为 Run）：

1. `add` 目标 paragraph，Run `query paragraph json` 取 `paraId`。
2. Run `raw /document`，按 `paraId` 抽出单个 `<w:p>` XML（不要用按行切分；`raw` 常为单行）。
3. 经宿主 **approval** 后 Run `raw-set`：`action replace`，用含 `lock=contentLocked` 的 `<w:sdt>…<w:sdtContent>[original w:p]</w:sdtContent></w:sdt>` 替换该 paragraph。**永不**把该 `raw-set` 写入 Batch JSON。

Verify with `query sdt json`，筛选 `.format.lock == "contentLocked"` 且 block type。Use only for legal attestations, compliance disclaimers, confidentiality clauses — regular intake fields do not need this.

### Role-gated fields (multi-role forms)

When one form is filled by two roles (patient vs physician; Party A vs Party B), use `lock=contentLocked` on the fields the other role must not touch. Under `protection=forms`, `contentLocked` SDTs display as read-only in Word; the intended role unprotects (or the admin swaps role-specific copies) to fill the other half.

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/form.docx",
    "operations": [
      {"command":"set","path":"/body/sdt[1]","props":{"lock":"sdtlocked"}},
      {"command":"set","path":"/body/sdt[2]","props":{"lock":"sdtlocked"}},
      {"command":"set","path":"/body/sdt[14]","props":{"lock":"contentLocked"}},
      {"command":"set","path":"/body/sdt[15]","props":{"lock":"contentLocked"}}
    ],
    "stop_on_error": true
  }
}
```

This is the core pattern for medical intake, two-party contracts, sequential-approval forms.

## Recipe — Contract / SOW template with MERGEFIELD + signature

Row-map across the three sub-recipes: SDT[1]=project_name, SDT[2]=contract_start, SDT[3]=payment_schedule, SDT[4]=signatory_name (inline). Run (sow-a) → (sow-b) → (sow-c) in order on the same file；each sub-recipe stays focused so a slip never cascades past one block.

### Recipe (sow-a) Boilerplate + cover + parties

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "create",
    "command_arguments": [
      "/workspace/sow.docx"
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
      "/workspace/sow.docx"
    ]
  }
}
```

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "set",
    "command_arguments": [
      "/workspace/sow.docx",
      "/",
      "--prop",
      "title=Statement of Work",
      "--prop",
      "docDefaults.font=Calibri",
      "--prop",
      "docDefaults.fontSize=12pt"
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
      "/workspace/sow.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Statement of Work",
      "--prop",
      "style=Heading1",
      "--prop",
      "size=20",
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
      "/workspace/sow.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=This Statement of Work ('SOW') is entered into between the parties identified below and governs the delivery of professional services.",
      "--prop",
      "size=11",
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
      "/workspace/sow.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Customer: "
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
      "/workspace/sow.docx",
      "/body/p[last()]",
      "--type",
      "field",
      "--prop",
      "fieldType=mergefield",
      "--prop",
      "name=CustomerName"
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
      "/workspace/sow.docx",
      "/body",
      "--type",
      "paragraph",
      "--prop",
      "text=Contract #: "
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
      "/workspace/sow.docx",
      "/body/p[last()]",
      "--type",
      "field",
      "--prop",
      "fieldType=mergefield",
      "--prop",
      "name=ContractNo"
    ]
  }
}
```

### Recipe (sow-b) SDT fields — all via Path A

```json
{
  "tool": "{{OFFICE_BATCH_TOOL}}",
  "arguments": {
    "file": "/workspace/sow.docx",
    "operations": [
      {"command":"add","parent":"/body","type":"sdt","props":{"type":"text","alias":"Project Name","tag":"project_name","text":"Enter project name"}},
      {"command":"add","parent":"/body","type":"sdt","props":{"type":"date","alias":"Contract Start Date","tag":"contract_start","format":"MM/dd/yyyy"}},
      {"command":"add","parent":"/body","type":"sdt","props":{"type":"dropdown","alias":"Payment Schedule","tag":"payment_schedule","items":"Full Prepayment,Net 30 Upon Delivery"}},
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Client Signature:","bold":true,"spaceBefore":"18pt","spaceAfter":"4pt"}},
      {"command":"add","parent":"/body","type":"sdt","props":{"type":"text","alias":"Signatory Name","tag":"signatory_name","text":"Authorized Signatory"},"after":"find:Client Signature:"}
    ],
    "stop_on_error": true
  }
}
```

# (Only reach for raw-set if a listItem's stored value must differ from its display text — see Path B.)

### Recipe (sow-c) Watermark + locks + document protection

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/sow.docx",
      "/",
      "--type",
      "watermark",
      "--prop",
      "text=CONFIDENTIAL",
      "--prop",
      "color=FF0000",
      "--prop",
      "rotation=315"
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
      "/workspace/sow.docx",
      "/body/sdt[1]",
      "--prop",
      "lock=sdtlocked"
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
      "/workspace/sow.docx",
      "/body/sdt[2]",
      "--prop",
      "lock=sdtlocked"
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
      "/workspace/sow.docx",
      "/body/sdt[3]",
      "--prop",
      "lock=sdtlocked"
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
      "/workspace/sow.docx",
      "/",
      "--prop",
      "protection=forms"
    ]
  }
}
```

先独立 Run `view forms` 复制 signatory_name path，再 Run `set` 其上 `lock=sdtlocked`（path 依赖，不要盲 Batch）。最后 Run `close`，再 Run `query field` — expect 2 MERGEFIELDs: CustomerName, ContractNo。

## Design principles (forms)

**Control-type decision tree:**

```
Date → type=date | Fixed list → type=dropdown | List + custom → type=combobox
Short text → type=text | Long text → type=richtext | Boolean → formfield checkbox
```

**Typography scale.** Spacing unit trap: `spaceBefore` / `spaceAfter` / `spaceLine` default to **twips** (1/20 pt) — always write `spaceBefore=18pt`.

| Element | Size | Style | Spacing |
|---|---|---|---|
| Form title (H1) | 20pt | Bold | `spaceBefore=0pt`, `spaceAfter=12pt` |
| Section heading (H2) | 14pt | Bold | `spaceBefore=18pt`, `spaceAfter=8pt` |
| Field label | 11pt | Bold | `spaceAfter=4pt` |
| Instructions / notes | 11pt | Italic `color=666666` | `spaceAfter=18pt` |

**Accessibility bump.** For medical / geriatric / accessibility-focused forms, raise field label + instruction to **12pt** (11pt default is tight for older users); keep section headings at 14pt.

**CJK forms:** set `docDefaults.font="Microsoft YaHei"` — Calibri lacks Chinese glyphs.

**Field ordering.** (1) Personal / ID, (2) role / classification, (3) dates, (4) supplemental free-text, (5) confirmation / signature.

**Yes/No + conditional follow-up** (common in compliance / medical intake): formfield checkbox followed by a richtext SDT whose `alias` carries the cue — e.g. `type=formfield` + `type=checkbox` + `name=has_cond` then `type=sdt` + `type=richtext` + `alias="If yes, explain"` + `tag=cond_detail`.

**Signature block order.** Label on its own paragraph, SDT on the next paragraph (with `spaceBefore=18pt` on the label, `spaceAfter=4pt` on the SDT). Never `Label: SDT` inline — Word renders the runs as touching, visually stuck together.

**Build order.** create+open → metadata → structure (headings, label paragraphs) → SDT/formfield skeletons (Path A 4 props) → Path B injections（逐步 Run + approval） → per-field lock → `protection=forms` LAST → close.

**Header / footer note.** Headers/footers are **predefined** when the section is created (default/first/even, 3 each). The first mutation must be `set` against the existing part, not `add` — `add … /header` returns `already exists` or silently no-ops. Inspect first with Run `query header json` to read the `type` values, then `set '/header[@type=default]'`. Only use `add` when creating an additional section with its own header/footer.

## Batch mode (brief)

For forms with many controls, Batch reduces overhead. Path A `add`/`set` 可同批；**Path B `raw-set` 必须独立 Run + approval，不得写入 Batch JSON。**

依赖链按顺序使用 Run：创建后先读取返回的稳定 path/identifier，并用实际返回值替换后续占位路径。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "add",
    "command_arguments": [
      "/workspace/form.docx",
      "/body",
      "--type",
      "sdt",
      "--prop",
      "type=text",
      "--prop",
      "alias=Full Name",
      "--prop",
      "tag=full_name",
      "--prop",
      "text=Enter name"
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
      "/workspace/form.docx",
      "/body",
      "--type",
      "sdt",
      "--prop",
      "type=dropdown",
      "--prop",
      "alias=Department",
      "--prop",
      "tag=dept",
      "--prop",
      "items=Engineering,Finance"
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
      "/workspace/form.docx",
      "/body/sdt[1]",
      "--prop",
      "lock=sdtlocked"
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
      "/workspace/form.docx",
      "/body/sdt[2]",
      "--prop",
      "lock=sdtlocked"
    ]
  }
}
```

然后独立 Run `set / prop protection=forms`。

- **P0 batch trap:** genuinely-unsupported props in batch are silently dropped, **no WARNING** (interactive `add` would print WARNING: UNSUPPORTED, exit 2). Defence: send only props Help `docx sdt` lists（`type/tag/alias/text/items/format/lock/...` are fine on `add`; `maxlength` is not）, and verify with a readback after the batch.
- Batch 支持 `add` / `set` / `remove` / `move` / `swap`（阶段 0 guidance）；readback / `validate` / `raw-set` 用独立 Run。

## Delivery Gate (executable)

Run every gate below after every form. Each gate must PASS. Any REJECT = do not deliver. **全部为独立 Run**（不要 bash / grep / jq）。Assumes document has been closed with Run `close`.

**Gate 1 — Validate.** Run `validate`。输出不得含 `[Schema]` error（`protection=forms` 不再产生 schema error）。否则 REJECT。

```json
{
  "tool": "{{OFFICE_RUN_TOOL}}",
  "arguments": {
    "command_name": "validate",
    "command_arguments": [
      "/workspace/form.docx"
    ]
  }
}
```

**Gate 2 — Token / placeholder leak.** Run `view text`，扫描 `___`（≥3 下划线）、`TBD`、`(fill in)`、`{{`、`xxxx`、`lorem`、`placeholder`。任一命中 → REJECT。

**Gate 3 — At least one structured field.** 分别 Run：
- `query sdt`（可加 `json`）
- `query formfield`
- `query field`

三者 `.data.results` 长度之和须 > 0；否则 REJECT（不是 form）。

**Gate 4 — Every SDT has alias + tag.** Run `query sdt json`。NOTE: prop 在 `.data.results[N].format.{prop}` — 用 `.format.alias` / `.format.tag`，never bare `.alias`。任一缺失 → REJECT。

**Gate 5 — Protection enforced + per-field lock inventory.** Run `get / json`；`.data.results[0].format.protection` 须为 `forms`。再 Run `view forms` 做 visual spot-check：every dropdown shows `items=`；every date shows `format=`；every locked SDT shows `lock=`。

**Gate 6 — No type=checkbox leaked onto any SDT.** Run `query sdt json`；不得出现 `.format.type == "checkbox"`。否则 REJECT（checkbox 只用 formfield）。

**Why `view issues` is not a gate.** It runs only prose-style checks (first-line-indent, heading size) and flags every form label as `Body paragraph missing first-line indent` — a false-positive avalanche on forms. Ignore for this skill. Use `validate` (schema integrity) and `view forms` (field inventory).

## Known Issues

| # | Issue | Behavior | Workaround |
|---|---|---|---|
| K1 | SDT `type=checkbox` not implemented | `add … type=sdt` + `type=checkbox` → Error, exit 1 | Use `type=formfield` + `type=checkbox` |
| K3 | SDT `maxlength` UNSUPPORTED on `add` | WARNING: UNSUPPORTED; element still created. (`items` / `format` / `lock` / `placeholderText` ARE supported on `add`; `name` is accepted but a no-op on SDT — use `alias`/`tag`) | Enforce length downstream; use `text` for initial content |
| K4 | SDT `items` / `format` / `type` not settable AFTER creation | `set prop items=…` → UNSUPPORTED (use raw-set instead). They ARE settable on `add` | Set at `add` time; to change later, Path B raw-set（Run + approval） or `remove` + re-add |
| K5 | FormField `maxlength` UNSUPPORTED | WARNING; formfield created | Enforce length in downstream validation |
| K6 | FormField dropdown `items` UNSUPPORTED | Dropdown formfield created with empty option list | Use an SDT dropdown with `items=` instead |
| K7 | Watermark `width` / `height` not settable | Watermark created without them; `opacity` IS settable on `add` | For an exact size, open Word + adjust the shape (Phase 2) |
| K9 | Batch mode silently drops genuinely-UNSUPPORTED props | No WARNING; batch reports success even when a prop was dropped | Keep batch SDT entries to props Help lists; verify with readback |
| K13 | FormField `name` > 20 characters | `add` exits 0; `validate` later reports MaxLength=20 | Keep `name` ≤ 20. SDT `alias` / `tag` have no such limit |
| K14 | `shd.fill` on a paragraph emits schema-invalid `<w:pPr>/<w:shd>` | `validate` reports schema errors; Word may still render | Apply highlight on the run (`shading=HEX`), or raw-set into run `<w:rPr>`（Run + approval） |
| K15 | `view forms` does NOT list MERGEFIELDs | Only SDT + formfield | Treat `query field` and `view forms` as two disjoint inventories |
| K16 | Header / footer are predefined at section creation | `add /header` returns already exists or no-ops | First mutation uses `set` against existing part after `query header` |
| K17 | Watermark injected into header may emit schema-invalid `<w:noProof>` | Extra `[Schema]` error under header SDT | After watermark add, Run raw-set remove `//w:noProof` per header part（approval；不进 Batch） |
| K18 | `query`/`get json` wrap props under `.data.results[N].format.{prop}` | Bare `.alias` / `.tag` / `.protection` returns null and Gate 4/5 falsely fail | Use `.data.results[].format.alias` / `.format.tag`; for `get /` use `.data.results[0].format.protection` |
| K19 | LibreOffice renders formfield checkbox as `☐☐` in PDF | Cosmetic; Word/WPS single box | Attribute to [RENDERER-BUG], not a form-quality defect |

## Phase 2 — enhance in Word

Some polish is out of CLI scope. Hand the file to a human for these; none are required for a valid form.

| Need | Why open Word |
|---|---|
| Real SDT checkbox with specific locking | `type=checkbox` exits 1; use Developer → Check Box Content Control (or a legacy FormField checkbox via host tools) |
| Prompt text ("Click here to enter a date") | Needs `placeholderDocPart` in `/word/glossary/document.xml` |
| Custom richtext default appearance | Adjust the referenced style in Word's style pane |
| Watermark resize | `width` / `height` not settable; drag shape handles |

(`picture` and `group` SDTs add directly via `type=sdt` + `type=picture` / `type=group` — no longer Phase-2 items.)

For the first four, build the skeleton once (Path C) and reuse.

## Help pointer

When in doubt: `{{OFFICE_HELP_TOOL}}` with `command_arguments` `docx` / element / verb+element. Help is the authoritative schema; this skill is the decision guide for building real fillable Word forms on top of it.
