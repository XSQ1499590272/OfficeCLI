# Word 操作测试用例清单

本文是 Word `.docx` 操作测试用例的候选清单，配合 `docs/test-case-reference-strategy.md` 使用。目标是先冻结当前代码行为，再支撑后续重构和迭代。

## 目标

- 固定当前 Word handler、CLI、batch、dump/readback 的可观察行为。
- 优先覆盖重构时最容易破坏的行为：路径、canonical key、alias、readback、关系重建、错误处理、失败不污染文件。
- 当当前行为和理想行为不一致时，先用测试记录当前行为；后续若要改行为，应显式调整测试和文档。
- 本文是总清单，不要求一次性全补完；先按“首批冻结基线用例”落地。

## 状态说明

| 状态 | 含义 |
|---|---|
| 待补 | 当前没有明确自动化测试覆盖，建议补充 |
| 已有单元覆盖 | 已有历史恢复测试覆盖一部分 handler 行为，但不是完整 CLI 覆盖 |
| 已有自动覆盖 | 已有新增自动化测试覆盖，可能是 unit、contract 或 integration |
| 部分自动覆盖 | 只有 handler/局部行为或单一 happy path 覆盖，仍缺 CLI、错误分支、round-trip 或跨进程验证 |
| CI smoke | 当前 CI 已覆盖最小链路，但断言较浅 |
| 示例待改造 | `examples/word` 已有可运行示例，适合抽取成集成测试 |
| 建议人工核验 | 自动断言成本较高，先用渲染或人工视觉核验兜底 |

## 测试落地目录建议

| 层级 | 建议位置 | 适合内容 |
|---|---|---|
| unit | `tests/OfficeCli.Tests/Unit/` | 直接调用 `WordHandler`、helper、selector、batch props parser，不启动 CLI 子进程 |
| contract | `tests/OfficeCli.Tests/Functional/` | schema/help 宣称的 add/set/get/query/readback 行为，使用临时 `.docx` 和 handler |
| integration | `tests/OfficeCli.Tests/Integration/` | 通过命令入口或构建产物跑 `create/add/set/get/query/remove/batch/dump/validate` |
| e2e | `tests/OfficeCli.Tests/E2E/` | 发布二进制、安装、resident、跨进程、平台相关行为；CI workflow 只负责调用这些测试 |
| visual | 手动验证记录或后续专门视觉测试 | screenshot/html/watch/diagram/textbox/chart 等高成本视觉稳定性 |

## 当前复审结论

当前清单和测试已经覆盖大部分 `WordHandler` 直接行为，尤其是 paragraph/run/table/media/sdt/revision/raw/dump 的 readback 和错误边界；真实 CLI 进程也已补入 `create/add/get/set/query/remove/validate` 最小闭环、resident 生命周期、发布产物 smoke、watch 的基础 mark/goto、示例脚本 smoke、schema/help、复杂 chart carrier 和 HTML preview contract。当前仍不能称为完整：现有不少条目是“一个测试覆盖一个功能族”，还没有把 schema 声明的每个操作面、每个命令的输入来源/输出模式、跨父级关系、失败原子性和跨进程协议拆成独立冻结点。字段实际刷新、复杂 package part/relationship 保真、全类型 CLI add/set 矩阵、watch/resident 并发与协议、发布包等价性仍需要继续细化。

补齐原则：handler 层已有覆盖的，不重复堆相同 happy path；优先补“源码里有独立命令面或 schema 声明，但当前测试只覆盖了内部 handler”的行为。下面新增的用例先进入清单并标明状态，后续实现测试时按这些 ID 逐项落地。

本轮整理后的规模统计：操作/场景条目共 364 条，其中已有单元覆盖 102 条、已有自动覆盖 97 条、部分自动覆盖 96 条、示例待改造 14 条、建议人工核验 2 条、CI smoke 1 条，新增待补 52 条；另有 154 条 `WORD-FREEZE-*` 冻结基线映射。这里的条目数不是 xUnit 方法数，也不是展开后的 test case 数，后续实现时一个条目可以对应一个或多个测试方法/数据行。

### 本轮复审后仍不全面的优先缺口

| 缺口 | 关联用例 | 为什么还要补 | 推荐下一步 |
|---|---|---|---|
| 字段实际刷新 | WORD-CMD-019/036, WORD-REF-024 | 已固定 refresh 的降级/成功 envelope、field fixture 的 backend 和 Pages 状态；TOC/PAGE/SEQ/REF 的实际 result/raw XML 差异仍未全量固定 | 扩展真实 field fixture，比较 refresh 前后 `raw/get/view issues`；没有 backend 时只固定降级，不伪造成功 |
| create 选项矩阵 | WORD-CMD-043 | `--type/--locale/--minimal` 的扩展名、默认字体/RTL、baseline parts 和 validate 已覆盖；Windows 发布包与更多 locale 仍待补 | 在 Windows 发布包重复标准、locale、minimal、无扩展名矩阵，补齐跨平台关键 parts/readback |
| get/view 选项矩阵 | WORD-CMD-044/045/046/047/057 | get depth、view 截断/过滤、HTML page/out、screenshot grid/native、pdf 成功或降级已各有 smoke；更多非法参数、page-count 和 resident lock 仍待补 | 继续补每个高成本分支的非法输入、resident lock 和 JSON/裸输出边界 |
| batch 输入与协议边界 | WORD-CMD-048/062 | `--commands`、`--input`、stdin、`--input -` 等价链路、空数组、UTF-8 BOM、JSONL 多行当前拒绝码、`--commands`/`--input` 互斥和坏输入不污染已固定；stdin 重定向 warning、全命令 stdout/stderr 纯度仍待补 | 继续补 stdin 重定向和全命令日志纯度，若未来支持 JSONL 再调整当前拒绝断言 |
| 已有文档保真 | WORD-DOC-024/025/026, WORD-MEDIA-025/027 | customXml/customXmlProperties/fontTable/webSettings、section refs、field cache 已有 fixture；主题、嵌入对象、未知 docProps、复杂 carrier 的完整保真矩阵仍可能丢失 | 使用带主题、字体、header/footer、chart/OLE 的 fixture，继续比较 parts、rels、content-types 和 validate |
| 失败原子性与锁竞争 | WORD-NEG-014/015/016 | failed set/raw-set/batch、resident file_locked/异常退出恢复、Unicode 路径已固定；remove/move/swap、busy 全命令、长路径和 Windows 分隔符仍待补 | 扩展失败快照到所有 mutation，并在 resident 占用、残留 marker、非 ASCII/长路径下固定错误和清理 |
| schema/help/runtime 漂移 | WORD-DOC-022/WORD-REV-014 | 单个功能已有很多测试，但 schema 声明、help 文案和 handler 实际支持面可能再次分叉 | 增加只读 contract，枚举 docx schema key 和 `officecli help docx` 输出，固定 add-only/read-only/settable 分类 |
| 复杂关系删除和 dump round-trip | WORD-MEDIA-025/027, WORD-NEG-013 | picture/OLE/chart 已有基础 round-trip，ActiveX/VML/chart externalData/header/footer 关系仍容易悬空 | 用小型 fixture 固定 rel/part 数量、validate 结果和 replay 后可查询性 |
| 示例脚本 smoke | WORD-DOC-023, WORD-TBL-006/011, WORD-MEDIA-005/011/013/015/017, WORD-REF-008/018, WORD-REV-007 | examples 是用户实际学习入口，但目前多数只是文档或脚本，不是回归测试 | 先抽每类 1 个短脚本跑临时目录，断言关键节点可 `query/get` |

### 本轮复审原则

上一轮清单已经能回答“哪些功能存在测试”，但还不能回答“某个功能在所有入口、输入、关系和失败状态下是否被固定”。本轮新增用例采用以下拆分规则：

- 同一功能至少区分 handler/helper、contract、真实 CLI、resident/跨进程和 round-trip 观察面；不能用 unit happy path 代替 CLI 协议测试。
- 同一命令至少区分成功、缺参/非法参、输入源、输出模式、重复执行和失败后的文件快照；不能只断言 exit code。
- 复杂 `.docx` 至少比较正文、目标 part、relationship、content type、binary payload、`validate` 结果和重新打开后的 readback；不能只比较 JSON 文本。
- 当前代码明确不支持的行为也进入清单，状态写成“待补”或“部分自动覆盖”，测试应固定“不支持”的错误和不变性，而不是修改源代码让测试通过。

## 单元测试候选

单元测试优先直接调用 handler/helper 或构造临时 OpenXML 文档，不启动 `officecli` 子进程。CLI 参数解析、真实发布包、安装和跨进程 resident 行为放到 integration/e2e。

落地原则：优先通过 `WordHandler` 公共方法和可观察 readback 固定行为；只有纯解析、格式化、路径索引、batch props 这类 helper 才直接测 helper。不要为了测试 private 实现而暴露新 API。

### 已有恢复覆盖与核心 readback

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-001 | unit | P0 | `tests/OfficeCli.Tests/UnitTest1.cs` | `WordHandler.Query("picture")` 识别 inline picture | 返回 1 个 picture；`wrap=inline`；路径可用于 `get` | 已有单元覆盖 |
| WORD-UNIT-002 | unit | P0 | `tests/OfficeCli.Tests/UnitTest1.cs` | `WordHandler.Query("picture")` 识别 anchored picture | `wrap/hPosition/vPosition/hRelative/vRelative` readback 正确 | 已有单元覆盖 |
| WORD-UNIT-003 | unit | P0 | `tests/OfficeCli.Tests/UnitTest1.cs` | `WordHandler.Query("ole")` 识别 OLE 并排除 picture | OLE 数量、`progId`、尺寸 readback 正确；picture 不混入 | 已有单元覆盖 |
| WORD-UNIT-004 | unit | P0 | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | SDT schema type values 与 add-time 实现一致 | 未实现 type 不出现在 schema；`type.set=false` | 已有单元覆盖 |
| WORD-UNIT-005 | unit | P0 | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | SDT canonical `type` add/get | `Add` 接受 `type`；`Get` 返回 `type`，不返回 `sdtType` | 已有单元覆盖 |
| WORD-UNIT-006 | unit | P0 | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | SDT `lock` 派生 `editable` | 4 种 lock 的 `editable` readback 正确；重开文件后保持 | 已有单元覆盖 |
| WORD-UNIT-007 | unit | P0 | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | checkbox SDT checked readback | `type=checkbox`；`checked=true` 可读 | 已有单元覆盖 |
| WORD-UNIT-008 | unit | P0 | `schemas/help/docx/paragraph.json` | `WordHandler.Add/Get/Set` 段落 canonical key | `align/spaceBefore/spaceAfter/lineSpacing` 写入和读取一致 | 已有单元覆盖 |
| WORD-UNIT-009 | unit | P0 | `schemas/help/docx/run.json` | run 格式写入和 readback | 文本、bold、italic、font、size、color、underline readback 正确 | 已有单元覆盖 |
| WORD-UNIT-010 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Selector.cs` | selector 解析和匹配 | `paragraph[style=...] > run[bold=true]`、`:contains()`、`:empty` 命中正确 | 已有单元覆盖 |
| WORD-UNIT-011 | unit | P1 | `schemas/help/docx/table-cell.json` | table/cell 合并的局部行为 | `gridspan/hmerge/vmerge` 后索引和被吸收 cell 行为正确；见 `tests/OfficeCli.Tests/Unit/WordTableTests.cs` | 已有单元覆盖 |
| WORD-UNIT-012 | unit | P1 | `schemas/help/docx/section.json` | section 长度和方向 readback | orientation、margin、columns 可读；`in`/`pt`/`dxa` 输入归一化为 cm readback 已固定；见 `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs` | 已有单元覆盖 |
| WORD-UNIT-013 | unit | P1 | `schemas/help/docx/bookmark.json` | bookmark/formfield 名称校验 | bookmark 重名/路径特殊字符当前保留；formfield 路径特殊字符和空白名拒绝 | 已有自动覆盖 |
| WORD-UNIT-014 | unit | P1 | `schemas/help/docx/revision.json` | revision marker 局部生成 | run scope ins/del/format revision marker 类型、author、id 正确；paragraph format、table/cell format、row ins/format 当前读回已固定；见 `tests/OfficeCli.Tests/Unit/WordRevisionTests.cs` | 已有单元覆盖 |
| WORD-UNIT-015 | unit | P1 | `schemas/help/docx/raw.json` | raw XML 写入失败不污染文档 | unknown part、XPath 无匹配、非法 XML fragment 抛出可解释错误且文档不变；见 `tests/OfficeCli.Tests/Unit/WordRawAndFindReplaceTests.cs` | 已有单元覆盖 |
| WORD-UNIT-016 | unit | P1 | `README.md` | batch item props 解析 | props `["key=value"]` 数组、JSON object 标量转换、malformed entry 跳过/错误 envelope 行为固定 | 已有自动覆盖 |
| WORD-UNIT-017 | unit | P1 | `README.md` | dump emitter 小型片段输出 | paragraph/table/picture 小片段生成的 batch item 形状和 replay 已固定 | 已有自动覆盖 |

### 路径、导航与选择器

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-018 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Navigation.cs` | positional path 导航 | `/body/p[1]`、`/body/p[1]/r[1]`、`/body/tbl[1]/tr[1]/tc[1]` 命中正确节点；见 `tests/OfficeCli.Tests/Unit/WordNavigationTests.cs` | 已有单元覆盖 |
| WORD-UNIT-019 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Navigation.cs` | stable id path 导航 | `p[@paraId=...]`、bookmark name 重开后命中；`sdt[@sdtId=...]` 已在 functional 固定 | 已有单元覆盖 |
| WORD-UNIT-020 | unit | P1 | `src/officecli/Core/PathIndex.cs` | path index 解析与格式化 | 1-based CLI 索引和数组索引互转稳定；当前仅做 `-1/+1`，不做非法索引校验 | 已有自动覆盖 |
| WORD-UNIT-021 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Selector.cs` | selector 属性等值和不等值 | `paragraph[align=...]`、`[align!=...]` 按当前规则过滤 | 已有自动覆盖 |
| WORD-UNIT-022 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Selector.cs` | selector child combinator | `paragraph[style=...] > run[bold=true]` 只返回满足父子条件的节点 | 已有自动覆盖 |
| WORD-UNIT-023 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Selector.cs` | selector pseudo | `:contains()`、`:empty`、`:no-alt` 命中当前行为；`:contains()`/`:empty` 已在 `tests/OfficeCli.Tests/Unit/WordSelectorTests.cs` 固定，`:no-alt` 已在 functional 固定 | 已有单元覆盖 |
| WORD-UNIT-024 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Query.cs` | query 后置过滤 | `paragraph[text~=...]` 当前直接查询不收窄；`type=paragraph` 返回 paragraph 节点 | 已有自动覆盖 |
| WORD-UNIT-025 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Navigation.cs` | header/footer 路径导航 | header/footer 内段落、run 路径命中 host part 内节点；图片路径已在 integration 固定；见 `tests/OfficeCli.Tests/Unit/WordHeaderFooterTests.cs` | 已有单元覆盖 |
| WORD-UNIT-026 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Navigation.DocSettings.cs` | document settings path | `/settings`、`/docDefaults` 路径读写命中正确 part；compatibility.mode 与 compatibility flag/preset readback 固定；见 `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs` | 已有单元覆盖 |

### 文本、段落、样式与文档设置

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-027 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Add.Text.cs` | add paragraph 基础文本 | 添加后 body 子节点顺序正确；文本和 paraId 存在；见 `tests/OfficeCli.Tests/Unit/WordBodyMutationTests.cs` | 已有单元覆盖 |
| WORD-UNIT-028 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Helpers.RunFormat.cs` | run 格式应用 | bold/italic/underline/color/font/size 写入到 rPr 并可 readback | 已有自动覆盖 |
| WORD-UNIT-029 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Set.Element.cs` | set run text | 修改文本不破坏已有 run 格式；空文本 set 后保留 run 节点与格式；见 `tests/OfficeCli.Tests/Unit/WordRunTests.cs` | 已有单元覆盖 |
| WORD-UNIT-030 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Set.Element.cs` | set paragraph props | align/spacing 写入 pPr 并 canonical readback；listStyle、rightIndent、hangingIndent readback 固定；firstLineIndent 与 hangingIndent 同设时当前不出现在 readback；见 `tests/OfficeCli.Tests/Unit/WordParagraphTests.cs` | 已有单元覆盖 |
| WORD-UNIT-031 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Helpers.Style.cs` | styleId/styleName 解析 | custom styleId、display name、空白文档实际存在的 `Normal` styleName 解析、未定义且含空格的 `Heading 1` 当前跳过行为已固定；见 `tests/OfficeCli.Tests/Unit/WordStyleAndNumberingTests.cs` | 已有单元覆盖 |
| WORD-UNIT-032 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.StyleList.cs` | add/set/get style | paragraph/character/table style 的 id/name/type/basedOn/readback、重复 custom id 错误已固定；见 `tests/OfficeCli.Tests/Unit/WordStyleAndNumberingTests.cs` | 已有单元覆盖 |
| WORD-UNIT-033 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Set.DocDefaults.cs` | docDefaults | 默认字体、字号、bold 写入并影响 direct `/docDefaults` 与 effective readback；RTL、alignment、spacing 的 `/docDefaults` readback 已固定；当前 `docDefaults.lang.*` 非 setter 支持面，后续若开放需补；见 `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs` | 已有单元覆盖 |
| WORD-UNIT-034 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Set.DocSettings.cs` | doc settings | root 与 `/settings` 路径的 docGrid、charSpacingControl、compatibility.mode、compatibility flag/preset 已固定；见 `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs` | 已有单元覆盖 |
| WORD-UNIT-035 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Set.SectionLayout.cs` | section layout | margin、orientation、columns、page size、direction、rtlGutter readback 固定；见 `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs` | 已有单元覆盖 |
| WORD-UNIT-036 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Helpers.FindReplace.cs` | find/replace 局部文本 | 普通 find、regex find、replace、bare find 错误、普通跨 run replace、跨 hyperlink 边界 replace 拒绝且不变更文本已固定；见 `tests/OfficeCli.Tests/Unit/WordRawAndFindReplaceTests.cs` | 已有单元覆盖 |
| WORD-UNIT-037 | unit | P2 | `src/officecli/Handlers/Word/WordHandler.I18n.cs` | i18n/RTL 属性 | run 级 lang.latin/ea/cs、rtl、complex-script 字体/字号/bold/italic readback 固定；RTL locale 写入 `lang.cs/locale`，新段落从 section 继承 `effective.direction/effective.rtl` 已固定；见 `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs` | 已有单元覆盖 |

### 表格、列表与编号

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-038 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Add.Table.cs` | add table | rows/cols/gridCol/cell paragraph 初始化结构固定；见 `tests/OfficeCli.Tests/Unit/WordTableTests.cs` | 已有单元覆盖 |
| WORD-UNIT-039 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Set.Element.cs` | set cell text/format | cell text、shd/fill、align、valign、padding readback 固定；见 `tests/OfficeCli.Tests/Unit/WordTableTests.cs` | 已有单元覆盖 |
| WORD-UNIT-040 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Set.Element.cs` | gridspan/hmerge | 合并后被吸收 cell 删除、后续 tc 索引变化固定；见 `tests/OfficeCli.Tests/Unit/WordTableTests.cs` | 已有单元覆盖 |
| WORD-UNIT-041 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Set.Element.cs` | vmerge | restart/continue 写入规则和 readback 固定；continue 见 `tests/OfficeCli.Tests/Unit/WordTableTests.cs` | 已有单元覆盖 |
| WORD-UNIT-042 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Helpers.SectionTable.cs` | table border/layout | border shorthand/sub-property、layout、cellSpacing、indent、colWidths 写入固定 | 已有自动覆盖 |
| WORD-UNIT-043 | unit | P1 | `schemas/help/docx/table-row.json` | row add/set/remove | row height/header/readback 和删除后索引行为固定 | 已有自动覆盖 |
| WORD-UNIT-044 | unit | P1 | `schemas/help/docx/table-column.json` | virtual column add/remove | gridCol 与每行 cell 同步；列删除后索引固定；见 `tests/OfficeCli.Tests/Unit/WordTableTests.cs` | 已有单元覆盖 |
| WORD-UNIT-045 | unit | P0 | `schemas/help/docx/numbering.json` | listStyle 高层编号 | bullet/ordered/none 当前 numId/abstractNum 行为和 listStyle/numFmt/numLevel readback 固定；见 `tests/OfficeCli.Tests/Unit/WordStyleAndNumberingTests.cs` | 已有单元覆盖 |
| WORD-UNIT-046 | unit | P1 | `schemas/help/docx/abstractNum.json` | abstractNum/level | 9 级默认 level、format/start/text/indent 写入固定；见 `tests/OfficeCli.Tests/Unit/WordStyleAndNumberingTests.cs` | 已有单元覆盖 |
| WORD-UNIT-047 | unit | P1 | `schemas/help/docx/num.json` | num instance | abstractNumId 引用、startOverride、非法引用行为固定；非法引用已在 functional 固定 | 已有单元覆盖 |

### 媒体、绘图、图表与关系

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-048 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Add.Media.cs` | add picture file/data URI | ImagePart、relId、docPr、尺寸、alt 写入固定；见 `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | 已有单元覆盖 |
| WORD-UNIT-049 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Add.Media.cs` | anchored picture props | wrap、h/v position、relative frame、behindText、h/v align 写入固定；基础 anchor 见 `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | 已有单元覆盖 |
| WORD-UNIT-050 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.ImageHelpers.cs` | picture crop/decorative/link | crop canonical readback、decorative 扩展、hyperlink rel 固定 | 已有自动覆盖 |
| WORD-UNIT-051 | unit | P0 | `src/officecli/Core/OleHelper.cs` | OLE data URI 和 content type | data URI decode、embedded package/legacy object content-type/fileSize readback 固定；见 `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | 已有单元覆盖 |
| WORD-UNIT-052 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Add.Media.cs` | add OLE | EmbeddedPackagePart、EmbeddedObjectPart、ProgID、display、名称、尺寸 readback 固定；见 `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | 已有单元覆盖 |
| WORD-UNIT-053 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Helpers.Chart.cs` | chart add/readback | chart type、title、series、categories、尺寸、value axis 基础 readback 固定；见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-UNIT-054 | unit | P1 | `src/officecli/Core/Chart/ChartHelper.cs` | chart series parser | 通过公共 chart readback 固定 inline values/categories、range-like input、颜色格式解析；helper 直接单测固定 chartType、range/category、series data、series color 当前解析行为；见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-UNIT-055 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Add.Media.cs` | inlined parts materialize | top-level chart part materialize、host relId 重写、child chart style part、host external rel、per-part external rel 重建已固定；见 `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | 已有单元覆盖 |
| WORD-UNIT-056 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Add.Diagram.cs` | diagram add | native mermaid 输入生成 group/textbox 结构、set/remove 和缺少 source 失败路径固定；见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-UNIT-057 | unit | P2 | `src/officecli/Handlers/Word/WordHandler.HtmlPreview.*.cs` | HTML preview basic render | 段落/表格最小 HTML、图片 data URI 和 alt 文本固定；见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |

### 引用、字段、表单与内容控件

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-058 | unit | P1 | `schemas/help/docx/hyperlink.json` | hyperlink external/internal | url rel、anchor、text、style readback 固定；见 `tests/OfficeCli.Tests/Unit/WordReferenceTests.cs` | 已有单元覆盖 |
| WORD-UNIT-059 | unit | P1 | `schemas/help/docx/bookmark.json` | bookmark pair | start/end 成对、id/name 生成和非法 name 校验固定；见 `tests/OfficeCli.Tests/Unit/WordReferenceTests.cs` | 已有单元覆盖 |
| WORD-UNIT-060 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Helpers.Field.cs` | field chain | begin/instr/separate/result/end 结构、query/get、`/field[N]` 虚拟路径直接 remove 被拒绝、`instrText`/`fieldChar` 结构 run 删除后的折叠 field 消失行为固定；见 `tests/OfficeCli.Tests/Unit/WordFieldAndFormTests.cs` | 已有单元覆盖 |
| WORD-UNIT-061 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.SeqEval.cs` | SEQ field evaluation | add SEQ field 的 body-order 编号缓存、`\\r N` reset 当前字段结果、ROMAN 格式、`recalcFields=seq` 按当前 instruction 重写缓存行为固定；基础 SEQ 见 `tests/OfficeCli.Tests/Unit/WordFieldAndFormTests.cs` | 已有单元覆盖 |
| WORD-UNIT-062 | unit | P1 | `schemas/help/docx/footnote.json` | footnote/endnote | 正文 reference 与 note part 内容成对，remove 行为固定；见 `tests/OfficeCli.Tests/Unit/WordNotesAndCommentsTests.cs` | 已有单元覆盖 |
| WORD-UNIT-063 | unit | P1 | `schemas/help/docx/comment.json` | comment range | comment range marker、CommentsPart 内容、author/date readback 固定；见 `tests/OfficeCli.Tests/Unit/WordNotesAndCommentsTests.cs` | 已有单元覆盖 |
| WORD-UNIT-064 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.FormFields.cs` | formfield | text/check/dropdown formfield 的 bookmark namespace 和 default state 固定；见 `tests/OfficeCli.Tests/Unit/WordFieldAndFormTests.cs` | 已有单元覆盖 |
| WORD-UNIT-065 | unit | P0 | `schemas/help/docx/sdt.json` | SDT per-type props | dropdown/combobox/date/picture/group/richtext add/get-only props 已固定；checkbox 已有历史覆盖；见 `tests/OfficeCli.Tests/Unit/WordSdtTests.cs` | 已有单元覆盖 |

### 分节内联标记与权限

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-076 | unit | P1 | `schemas/help/docx/pagebreak.json` | page/column break | add pagebreak 后 query/get 返回 `breakType` 和稳定路径；见 `tests/OfficeCli.Tests/Unit/WordBreakTabPermissionTests.cs` | 已有单元覆盖 |
| WORD-UNIT-077 | unit | P1 | `schemas/help/docx/tab.json`, `schemas/help/docx/ptab.json` | tab stop 与 positional tab | tab add/set/remove 当前会留下空 `w:tabs` schema 错误；ptab add/set/query 无 validate 错误；见 `tests/OfficeCli.Tests/Unit/WordBreakTabPermissionTests.cs` | 已有单元覆盖 |
| WORD-UNIT-078 | unit | P2 | `schemas/help/docx/permStart.json` | editing permission range marker | permStart/permEnd 成对属性、删除 marker 不删文本、非法 id 不污染文档；见 `tests/OfficeCli.Tests/Unit/WordBreakTabPermissionTests.cs` | 已有单元覆盖 |

### 基础 mutation 与错误边界

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-079 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Query.cs`, `src/officecli/Handlers/Word/WordHandler.Mutations.cs` | query/remove 基础行为 | query `:contains()` 返回稳定路径；remove 删除目标段落但保留兄弟节点；见 `tests/OfficeCli.Tests/Unit/WordBodyMutationTests.cs` | 已有单元覆盖 |
| WORD-UNIT-080 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Mutations.cs` | move/copy 基础行为 | move 重新排序 body 段落且不丢文本；copy 克隆段落并保留文本；见 `tests/OfficeCli.Tests/Unit/WordBodyMutationTests.cs` | 已有单元覆盖 |
| WORD-UNIT-081 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Add.cs`, `src/officecli/Handlers/Word/WordHandler.Set.Element.cs` | 失败不污染文档 | invalid parent、picture 缺 src/非法尺寸、unsupported set、损坏 docx 当前错误行为不写入半成品；见 `tests/OfficeCli.Tests/Unit/WordErrorBoundaryTests.cs` | 已有单元覆盖 |

### 复杂对象基础单元

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-082 | unit | P1 | `schemas/help/docx/equation.json` | equation add/set/remove | display equation 返回 `/body/oMathPara[N]`，inline equation 返回 `/oMath[N]`，set formula 更新文本，remove 后 query 为空；见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 已有单元覆盖 |
| WORD-UNIT-083 | unit | P1 | `schemas/help/docx/textbox.json` | textbox add/set/get | textbox 内容树可通过 `/body/textbox[N]/p/r` 寻址；shape surface 可 set，text/position add-only 当前行为固定；见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 已有单元覆盖 |
| WORD-UNIT-084 | unit | P2 | `schemas/help/docx/shape.json` | floating shape add/set/remove | shape raw drawing tree、geometry/fill/size 子路径读回；set legacy alias 和 remove 行为固定；见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 已有单元覆盖 |
| WORD-UNIT-085 | unit | P1 | `schemas/help/docx/watermark.json` | watermark text/image 当前行为 | text watermark add/set/remove、非法 rotation、image 当前退回默认 text watermark 行为固定；见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 已有单元覆盖 |
| WORD-UNIT-086 | unit | P1 | `schemas/help/docx/chart.json` | chart add/get/query 基础面 | `Add("chart")` 返回 `/chart[N]`；`get/query` 可读 chartType/title/categories/series/size；`validate` 为空；见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-UNIT-087 | unit | P1 | `schemas/help/docx/chart-series.json`, `schemas/help/docx/chart-axis.json` | chart series/axis set | series name/values、range-backed refs/color、value axis title/min/max/format readback；非法 series 数值当前会清空 values 并留下 schema 错误；见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-UNIT-088 | unit | P1 | `schemas/help/docx/diagram.json` | native diagram 当前边界 | native mermaid 生成 group 和首个 textbox；group set/remove；缺少 source 不创建半成品；见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-UNIT-089 | unit | P2 | `src/officecli/Handlers/Word/WordHandler.HtmlPreview.*.cs` | HTML preview 当前输出 | 段落、表格、图片 data URI、alt 文本出现在 HTML 中；见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-UNIT-090 | unit | P1 | `schemas/help/docx/run.json`, `src/officecli/Handlers/Word/WordHandler.Helpers.RunFormat.cs` | run 高级开关格式 | `highlight/strike/dstrike/caps/smallcaps/vanish/outline/shadow/emboss/imprint/noProof` 写入、清空和 readback 当前行为固定；见 `tests/OfficeCli.Tests/Unit/WordRunTests.cs` | 已有单元覆盖 |
| WORD-UNIT-091 | unit | P1 | `schemas/help/docx/run.json`, `src/officecli/Handlers/Word/WordHandler.Helpers.TextEffect.cs` | w14 文字效果 | `outline/fill/w14shadow/w14glow/w14reflection` 解析、OOXML 节点、readback 字符串和非法 spec 报错固定；见 `tests/OfficeCli.Tests/Unit/WordRunTests.cs` | 已有单元覆盖 |
| WORD-UNIT-092 | unit | P1 | `schemas/help/docx/run.json` | run shading 与 theme color | `fill/shading/shd` solid 与 pattern readback 分流；`color` theme-linked compound form round-trip 保留 theme linkage；见 `tests/OfficeCli.Tests/Unit/WordRunTests.cs` | 已有单元覆盖 |
| WORD-UNIT-093 | unit | P1 | `schemas/help/docx/run.json` | range offset 格式化 | `range=START:END` 拆分文本并只格式化半开区间；越界、`find`+`range`、`text`+`range` 错误、多段逗号 range 的排序和跨段落格式化均已固定；见 `tests/OfficeCli.Tests/Unit/WordRunTests.cs`、`tests/OfficeCli.Tests/Functional/WordFindReplaceContractTests.cs` | 已有自动覆盖 |
| WORD-UNIT-094 | unit | P1 | `src/officecli/Handlers/WordHandler.cs` | 自动补齐稳定 ID | editable 打开后补齐缺失 `paraId/docPropId/sdtId/bookmarkId`；read-only 打开不写回；重复 ID 当前修正规则固定；见 `tests/OfficeCli.Tests/Unit/WordStableIdTests.cs` | 已有单元覆盖 |
| WORD-UNIT-095 | unit | P1 | `src/officecli/Handlers/WordHandler.cs`, `src/officecli/Core/RawXmlHelper.cs` | strict attribute sanitizer 与 repair path | strict namespace 属性清理、打开失败释放文件句柄已固定；zip URI raw 读取和 repair retry 待补；见 `tests/OfficeCli.Tests/Unit/WordStrictAttributeSanitizerTests.cs` | 部分自动覆盖 |
| WORD-UNIT-096 | unit | P1 | `src/officecli/Handlers/WordHandler.cs` | binary payload extraction | picture/OLE/media `TryExtractBinary` 成功返回 bytes/contentType；无 payload 的 paragraph/run 返回 false 且不创建空文件；见 `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | 已有单元覆盖 |
| WORD-UNIT-097 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Add.Text.cs`, `src/officecli/Core/TextEscape.cs` | 文本输入特殊处理 | `text` 中真实换行和 tab 生成 `<w:br>/<w:tab>`；`{page}/{pages}` 生成 PAGE/NUMPAGES field；非法 XML 控制字符被拒绝且不污染文档；C-style escape 的支持、未知/尾部反斜杠和 null/empty 行为已固定；见 `tests/OfficeCli.Tests/Unit/WordTextInputBoundaryTests.cs` | 已有单元覆盖 |
| WORD-UNIT-098 | unit | P1 | `schemas/help/docx/tab.json`, `schemas/help/docx/ptab.json` | tab/ptab 扩展边界 | tab 允许负 position；非法 val/leader 错误清晰；ptab 可携带 run 格式并在 header/footer 场景 readback；见 `tests/OfficeCli.Tests/Unit/WordBreakTabPermissionTests.cs` | 已有单元覆盖 |

### 修订、raw、batch 与 dump

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-066 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Set.Revision.cs` | run revision marker | ins/del/format marker 和 text/query readback 固定；run-level tracked move 生成 moveFrom/moveTo 共享 id/author 并可 query；见 `tests/OfficeCli.Tests/Unit/WordRevisionTests.cs` | 已有单元覆盖 |
| WORD-UNIT-067 | unit | P0 | `src/officecli/Handlers/Word/WordHandler.Set.Revision.cs` | paragraph/table revision marker | paragraph format、paragraph ins 拒绝、table/cell format、row ins、row format/trPrChange 当前结构已通过 unit 固定；见 `tests/OfficeCli.Tests/Unit/WordRevisionTests.cs` | 已有单元覆盖 |
| WORD-UNIT-068 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.Set.Revision.cs` | accept/reject revision | run insertion/deletion/format/move accept/reject 当前结果已通过 contract 固定 | 已有自动覆盖 |
| WORD-UNIT-069 | unit | P1 | `src/officecli/Core/RawXmlHelper.cs` | raw XML helper | XPath 命中、命名空间、写属性/节点、非法 XML 行为固定；Word handler raw/raw-set 见 `tests/OfficeCli.Tests/Unit/WordRawAndFindReplaceTests.cs` | 已有单元覆盖 |
| WORD-UNIT-070 | unit | P1 | `src/officecli/BatchTypes.cs` | batch props 解析 | array props、object props、malformed array entry 当前宽容/错误规则固定；裸 string props 不支持 | 已有自动覆盖 |
| WORD-UNIT-071 | unit | P1 | `src/officecli/Core/BatchExecutor.cs` | batch 执行策略 | 默认继续、stop-on-error、单项错误 envelope 当前行为固定 | 已有自动覆盖 |
| WORD-UNIT-072 | unit | P1 | `src/officecli/Handlers/Word/WordBatchEmitter.cs` | dump emitter root | `/body` 子树 dump item target path 和 replay 已固定；完整 document dump 的 numbering/styles/docDefaults/theme/settings 等关键资源项在 body item 之前输出已由 `tests/OfficeCli.Tests/Unit/WordDumpEmitterTests.cs` 固定 | 已有单元覆盖 |
| WORD-UNIT-073 | unit | P1 | `src/officecli/Handlers/Word/WordBatchEmitter.Paragraph.cs` | paragraph dump emitter | paragraph text item、显式 run/hyperlink/field、textbox add item 与内部文本 set 输出已由 `tests/OfficeCli.Tests/Unit/WordDumpEmitterTests.cs` 固定 | 已有单元覆盖 |
| WORD-UNIT-074 | unit | P1 | `src/officecli/Handlers/Word/WordBatchEmitter.Table.cs` | table dump emitter | table rows/cols、cell text set item、gridspan/colspan、vmerge 输出已固定；见 `tests/OfficeCli.Tests/Unit/WordDumpEmitterTests.cs` | 已有单元覆盖 |
| WORD-UNIT-075 | unit | P1 | `src/officecli/Handlers/Word/WordBatchEmitter.AuxParts.cs` | aux part dump emitter | picture data URI/name/alt 输出和 replay 已固定；full dump 中 chart/OLE 不产生 unsupported warning 且发出 `add chart`/`add ole` 已固定；inlinedparts complex carrier 的 part/child/per-part external props 已由 `tests/OfficeCli.Tests/Unit/WordDumpEmitterTests.cs` 固定 | 已有单元覆盖 |

### 本轮复审新增的单元细化项

下面这些用例不是把已有大类换个名字，而是把当前仍可能在重构中漂移的局部契约拆开。它们先进入清单，状态为“待补”；实现时优先使用现有 `WordTestBase` 和公共 handler API，不为测试暴露 private API。

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-099 | unit | P0 | `schemas/help/docx/*.json`, `SchemaHelpLoader` | docx schema operation matrix | 已固定全部嵌入 docx schema 的 operations 标记、属性 type/操作标记结构，以及 `word`、`/`、`p`、`tbl`、`tc` 别名解析；schema 与 handler 全量可达操作 parity 仍待补；见 `tests/OfficeCli.Tests/Unit/WordSchemaParityTests.cs` | 已有自动覆盖 |
| WORD-UNIT-100 | unit | P0 | `TypedAttributeFallback`, `ParsePropsArray`, `PathAliases` | typed property parser matrix | dotted fallback 已固定元素别名大小写、单层颜色 `#` 归一化、SDK 已知/未知属性、nested existing-only 行为、font alias merge、malformed key 和 PathAliases；bool/int/length/color/font-size/enum 的完整大小写、空白、非法值和 `key=value`/JSON object 等价矩阵仍待补；见 `tests/OfficeCli.Tests/Unit/WordTypedAttributeFallbackTests.cs`、`WordPathAliasTests.cs` | 部分自动覆盖 |
| WORD-UNIT-101 | unit | P0 | `WordHandler.Navigation`, `WordHandler.Selector` | path/selector grammar boundary | 正常 singleton、1-based index、stable id、`:contains`、`:empty`、`>`、`or`、`@` 属性前缀和 bracket 内 `>` 比较符已有覆盖；未闭合 selector/path 当前抛出可解释 `ArgumentException`，不静默返回空结果已固定；带引号属性值、空 selector 和更多错误码矩阵仍待补；见 `tests/OfficeCli.Tests/Unit/WordSelectorTests.cs`、`WordNavigationTests.cs`、`WordPathAliasTests.cs` | 部分自动覆盖 |
| WORD-UNIT-102 | unit | P0 | `WordHandler.Add`, `InsertPosition` | insert position boundary | `InsertPosition` 的 append/index/before/after、负数/越界原值和 anchor 异常传播已由 unit 固定；真实 handler 的 parent/ID 不变性和跨 parent 仍待补；见 `tests/OfficeCli.Tests/Unit/WordMutationPositionTests.cs` | 部分自动覆盖 |
| WORD-UNIT-103 | unit | P0 | `WordHandler.Mutations`, `WordHandler.ImageHelpers` | relationship host and relId collision | 同一 host part 中重复添加/复制 picture、OLE、hyperlink、chart 时 relId 唯一；删除一个引用不误删共享或其他 host 的 part；validate 结果稳定 | 待补 |
| WORD-UNIT-104 | unit | P1 | `WordHandler`, `WordBatchEmitter.Resources` | package content-type and auxiliary-part lifecycle | theme/styles/numbering/settings/core/app/customXml/customXmlProperties/fontTable/webSettings 等 part 的创建、复用、删除、content type 和 relationship 类型固定 | 待补 |
| WORD-UNIT-105 | unit | P1 | `WordHandler.Helpers.Field`, `WordHandler.SeqEval` | field instruction and cached-result parser | 已固定 dynamic/interactive field 分类、SEQ next/repeat/reset、ROMAN/alphabetic 缓存重算，以及 `\#` unsupported switch 保留既有 cache；PAGE/NUMPAGES/REF/PAGEREF/TOC 的大小写、空 result、未知 switch、重复 identifier 和更多 field parser 分支仍待补；见 `tests/OfficeCli.Tests/Unit/WordFieldParserTests.cs`、`tests/OfficeCli.Tests/Functional/WordFieldContractTests.cs` | 部分自动覆盖 |
| WORD-UNIT-106 | unit | P1 | `WordHandler.Helpers.HeaderFooter`, `WordHandler.Set.SectionLayout` | header/footer linkage | 已固定 default/first/even reference、`titlePg`、`evenAndOddHeaders`、重复 type 拒绝和未知 type 不创建 part；link-to-previous 继承/解除、删除最后一个 part 和跨 section linkage 仍待补；见 `tests/OfficeCli.Tests/Unit/WordHeaderFooterTests.cs`、`tests/OfficeCli.Tests/Functional/WordSectionRelationshipTests.cs` | 部分自动覆盖 |
| WORD-UNIT-107 | unit | P1 | `WordHandler.Add.Structure`, `WordHandler.Set.SectionLayout` | section mutation matrix | 已固定 section break 的 nextPage/continuous/evenPage/oddPage/nextColumn、page size/margin/orientation/columns/direction、page numbering、line numbering、column separator 和非法 break type；首/尾 section 增删、body `sectPr` 归属、更多 ID/path 稳定性仍待补；见 `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs`、`tests/OfficeCli.Tests/Functional/WordSectionRelationshipTests.cs` | 部分自动覆盖 |
| WORD-UNIT-108 | unit | P1 | `WordHandler.Set.Element`, table schemas | table row/cell property matrix | 已固定 row height/height.rule/header/cantSplit/hidden/cellSpacing/rowAlign/gridBefore/wBefore，以及 cell width 百分比、tcFitText、nowrap、textDirection、shading/padding；table layout/border/spacing/indent 已有局部覆盖，更多 add/set/remove 默认值和非法组合仍待补；见 `tests/OfficeCli.Tests/Unit/WordTableTests.cs`、`tests/OfficeCli.Tests/Functional/WordTableContractTests.cs`、`WordTableRowContractTests.cs` | 部分自动覆盖 |
| WORD-UNIT-109 | unit | P1 | `WordHandler.Add.Media`, `WordHandler.ImageHelpers`, `OleHelper` | media optional property matrix | 已固定 picture crop、decorative/hidden、click link、wrap/position/relativeHeight，以及 OLE display/name/size/content type 基础组合；rotation/flip、crop 清空、更多 wrap polygon/distance 和非法范围仍待补；见 `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs`、`tests/OfficeCli.Tests/Functional/WordPictureContractTests.cs`、`WordOleContractTests.cs` | 部分自动覆盖 |
| WORD-UNIT-110 | unit | P1 | `WordHandler.Add.Diagram`, `WordHandler.Add.Media`, complex builders | complex object validation boundary | chart/diagram/equation/shape/textbox/watermark 的最小输入、缺 source、坏 data、未知 property、部分 child 失败均不留下半成品或 dangling relationship | 待补 |
| WORD-UNIT-111 | unit | P1 | `WordHandler.Helpers.FindReplace`, `WordHandler.FormFields` | token and placeholder boundary | 普通文本、跨 run、跨 hyperlink、跨 field、跨 note part、重复 key、转义花括号、空 replacement 的替换范围和未替换边界固定 | 待补 |
| WORD-UNIT-112 | unit | P1 | `WordBatchEmitter`, `BatchExecutor` | dump determinism and warning ordering | 已固定同一文档连续 full dump 的 item 数量、资源/正文顺序和 JSON 序列化稳定；stable ID 全量矩阵、warning 去重、复杂 carrier 顺序和不支持元素 warning 仍待补；见 `tests/OfficeCli.Tests/Unit/WordDumpEmitterTests.cs`、`WordDumpFilterTests.cs` | 部分自动覆盖 |
| WORD-UNIT-113 | unit | P2 | `WordHandler.HtmlPreview.*` | HTML escaping and data-path safety | 已固定共享 HTML 文本转义、CJK/RTL 不变形，以及 PNG/JPEG 与 WMF/EMF/TIFF data URI 的降级边界；Word 专用连续空格、attribute/path 转义和脚本片段安全仍待补；见 `tests/OfficeCli.Tests/Unit/WordHtmlSafetyTests.cs` | 部分自动覆盖 |
| WORD-UNIT-114 | unit | P1 | `WordHandler.Validate`, OpenXML validator adapter | validation diagnostic normalization | 错误的 `type/description/part/path` 映射、多个错误排序、warning 与 error 分离、空结果和重复诊断去重固定 | 待补 |
| WORD-UNIT-115 | unit | P1 | `ResidentClient`, `ResidentServer`, `ResidentFlushPolicy` | resident wire and at-most-once contract | request/response JSON 字段、空响应、超长消息、连接失败、重试只发生在 connect 阶段；mutation 不因 read failure 被重复发送 | 待补 |
| WORD-UNIT-116 | unit | P1 | `RawXmlHelper`, `WordHandler` | raw URI and repair lifecycle | semantic path、zip URI 的 query/fragment、namespace XPath、strict attribute repair、打开失败后的 stream/file handle 释放和二次打开行为固定 | 待补 |

## 基础命令与端到端链路

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-CMD-001 | e2e | P0 | `.github/workflows/build.yml` | `create -> add paragraph -> get -> close` 最小 `.docx` 链路 | 产物存在；`get /body/p[1]` 返回预期文本；`close` 后磁盘可读 | CI smoke |
| WORD-CMD-002 | integration | P0 | `README.md` | `set` 修改段落或 run 属性 | `get` 返回 canonical key；`validate` 无错误 | 已有自动覆盖 |
| WORD-CMD-003 | integration | P0 | `README.md` | `query` 按元素类型和选择器检索 | 返回路径稳定；数量符合预期；不存在时为空结果 | 已有自动覆盖 |
| WORD-CMD-004 | integration | P0 | `README.md` | `remove` 删除段落、run、表格子元素 | 删除后 `query/get` 不再返回目标；其他节点路径可解释 | 已有自动覆盖 |
| WORD-CMD-005 | integration | P1 | `README.md` | `move` 调整 body 内元素顺序 | 顺序变化正确；内容和格式不丢失 | 已有自动覆盖 |
| WORD-CMD-006 | integration | P0 | `README.md` | `batch --commands` 重放多步 Word 修改 | 单命令结果与 batch 结果一致；错误项按策略返回 | 已有自动覆盖 |
| WORD-CMD-007 | integration | P1 | `README.md` | `dump existing.docx -> batch new.docx` 小型 round-trip | 新文档 `validate` 通过；关键段落、表格、图片可查询 | 已有自动覆盖 |
| WORD-CMD-008 | integration | P1 | `README.md` | `dump` 子树路径，例如 `/body/tbl[1]` | 表格子树 dump 后 batch 到新文档，目标存在且非目标段落不出现 | 已有自动覆盖 |
| WORD-CMD-009 | integration | P0 | `README.md` | `validate` 检查生成后的 `.docx` | 空文档成功结果，以及空 chart part 的 CLI `success=false/data.errors/count/type/description/part` 结构已固定；更多 dangling relationship 和错误类别待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-010 | e2e | P1 | `README.md` | `view text/annotated/outline/stats/issues --json` | 真实 CLI 子进程输出 JSON 可解析；关键文本和各 mode 的稳定结构键存在 | 已有自动覆盖 |
| WORD-CMD-011 | e2e | P2 | `README.md` | `view html/screenshot` 渲染 Word 文档 | HTML 非空；截图文件存在且非空白 | 建议人工核验 |
| WORD-CMD-012 | integration | P1 | `README.md` | `open -> set -> save -> 外部读取 -> close` resident flush 链路 | resident pipe 层已固定 mutation 先进内存、`save` 后外部 OpenXML 可读、`close` 后落盘；真实 CLI `open` 自举、CLI add/query 和 close 后外部读取已固定；见 `tests/OfficeCli.Tests/Integration/WordResidentLifecycleTests.cs`、`tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-013 | e2e | P1 | `README.md` | `OFFICECLI_RESIDENT_FLUSH=each` 每次修改落盘 | 真实 CLI resident 下 `add` 返回后无需 `save/close`，外部 OpenXML 读取已能看到磁盘内容；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-014 | e2e | P2 | `README.md` | `watch` 本地预览服务 | 服务启动；HTML 端点可访问；修改后版本刷新 | 建议人工核验 |
| WORD-CMD-015 | e2e | P1 | `README.md` | `merge template.docx out.docx data.json` | body/table/header/footer、对象/数组/literal key、未命中 key、`file_exists`、`invalid_json`，以及 footnote/endnote/comment 替换已由真实 CLI 固定；跨 run 占位符当前保留未解析也已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-016 | integration | P1 | `schemas/help/docx/raw.json` | `raw` 读取 document/settings/styles 等 part | XPath 命中；输出 XML/JSON 可解析；不存在 part 有结构化报错 | 已有自动覆盖 |
| WORD-CMD-017 | integration | P1 | `schemas/help/docx/raw.json` | `raw-set` 修改单个 XML 节点或属性 | 写回后 `raw` 可读到新值；`validate` 通过；错误 XPath 不污染文件 | 已有自动覆盖 |
| WORD-CMD-018 | integration | P1 | `schemas/help/docx/raw.json` | `add-part` 创建新 part 和 relationship | chart part 可创建、raw 可读、CLI 路径、重复 chart part 的不同 relId 和递增 path 已固定；main document relationship 类型由现有 unit/raw 验证覆盖；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-019 | e2e | P2 | `README.md` | `refresh` 更新 TOC / PAGE / cross-reference | 支持平台执行成功；不支持平台返回可解释降级；当前已固定 backend 成功消息与扩展属性页数写入，字段结果深度变化仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-020 | e2e | P1 | `plugins/plugin-protocol.md` | 插件 `dump-reader` 产出 docx batch 并由主程序回放 | 最小 dump-reader 的 `--info`、JSONL `BatchItem` replay 到 docx、生成文档可 query/validate，以及顶层 JSON 数组触发 `corrupt_batch` 已固定；真实第三方插件、Windows plugin host 和超时/heartbeat 矩阵待补；见 `tests/OfficeCli.Tests/Integration/WordPluginContractTests.cs` | 部分自动覆盖 |
| WORD-CMD-021 | e2e | P0 | `src/officecli/CommandBuilder.cs` | `create` CLI 行为 | 新建 docx 成功；已存在文件无 `--force` 失败；`--force` 覆盖；输出 JSON envelope 稳定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-022 | e2e | P0 | `src/officecli/CommandBuilder.Add.cs` | `add` CLI 参数面 | `--type/--prop/--index/--before/--after` 生效；`--from` 复制；`--from` 和 `--prop` 同用按当前错误拒绝；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-023 | e2e | P0 | `src/officecli/CommandBuilder.GetQuery.cs` | `get --json` 与 `get --save` | 单 path 返回 `{matches,results}`；picture/OLE 可保存 payload；无 payload 保存失败且不创建空文件；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-024 | e2e | P0 | `src/officecli/CommandBuilder.GetQuery.cs` | `query --find` 与 JSON children shape | `--find` 大小写宽容、matches 和 children array 最小形态、child combinator `paragraph > run[bold=true]`、boolean `or` selector 和 hydrated children 已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-025 | e2e | P0 | `src/officecli/CommandBuilder.Set.cs` | `set` CLI 参数与 selector set | path `set --prop align=center` 后 `get` readback、缺属性 `missing_property`、裸 `key=value` 的 `missing_prop_flag`、unsupported props warning/envelope 且已支持属性仍生效、`/body/p[text~=target]` scoped selector 多命中并逐项 readback 已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-026 | e2e | P1 | `src/officecli/CommandBuilder.Add.cs`, `src/officecli/CommandBuilder.Set.cs` | mutation selector guard | bare selector `set/remove` 拒绝为 `bare_selector_rejected`、slash path `/` scoped set、`/body/...` scoped selector 和 resident forward 路径均已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-027 | e2e | P1 | `src/officecli/CommandBuilder.Add.cs`, `src/officecli/Handlers/Word/WordHandler.Mutations.cs` | `remove/move/copy` CLI 面 | path remove 后 query 为空、`remove /body/p` 当前删除首个匹配段落、`move --before/--after/--index` 输出和顺序、`add --from --index` copy target parent 输出和 readback、`/body/ole[1]` indexed shorthand 删除与 validate、resident remove 转发及 close 后落盘已固定；resident move/copy 的内存顺序、显式 `save` 成功和 close 后磁盘顺序已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-028 | integration | P0 | `src/officecli/CommandBuilder.Batch.cs` | batch JSON envelope parity | non-resident 与 resident 的 `--json` envelope、`results/summary`、失败 item 回传、默认继续和 `--stop-on-error` 的 executed/skipped 结果均已固定；见 `tests/OfficeCli.Tests/Integration/WordResidentLifecycleTests.cs` | 已有自动覆盖 |
| WORD-CMD-029 | integration | P0 | `src/officecli/CommandBuilder.Batch.cs`, `src/officecli/CommandBuilder.Save.cs` | resident `open/save/close` 生命周期 | resident pipe 层、非 resident `save` no-op、真实 CLI `open` 自举及 `close` 后落盘均已固定；`OFFICECLI_RESIDENT_FLUSH=each/off/固定 1 秒/auto` 的单 mutation 磁盘可见性已固定，异常保存路径仍待补；见 `tests/OfficeCli.Tests/Integration/WordResidentLifecycleTests.cs`、`tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-030 | integration | P1 | `src/officecli/CommandBuilder.Batch.cs` | resident batch 特殊项 | resident 内 batch 遇到 `open/close` 跳过且返回成功项、会话保持存活、batch 内 query 可见新增内容、显式 `save` 后磁盘可见，以及 `each` 策略下 batch 返回即外部可见已固定；`off/固定间隔/adaptive auto` 的 batch 矩阵仍待补 | 部分自动覆盖 |
| WORD-CMD-031 | e2e | P1 | `src/officecli/CommandBuilder.Dump.cs` | `dump --out` 与 stdout/stderr 协议 | `--out file` 写裸 JSON array；`--out -` 输出 stdout；unsupported format 返回 `invalid_format`；orphan footnote warning 进入 JSON envelope 的 `warnings` 并写入 stderr，且不污染 batch data 已固定；其他 warning 类型待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-032 | e2e | P1 | `src/officecli/CommandBuilder.Raw.cs` | `raw/raw-set/add-part` CLI 协议 | 语义路径 `/document`、zip URI `/word/document.xml`（含 query/fragment）、不存在 part 的 `invalid_value`、`raw-set setattr`、错误 XPath 不污染、非法 XML 返回 `internal_error` 且原文档保持不变、`add-part chart` relId/path、`/chart[1]` raw 可读和空 chart 的 validate error envelope 已固定；复杂 chart validation 仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-033 | e2e | P2 | `src/officecli/CommandBuilder.Watch.cs`, `src/officecli/CommandBuilder.GetQuery.cs` | watch selection 与 `get selected` | 无 watch、运行 watch 但无 selection 时的 JSON envelope，以及 empty selection 已固定；多 selection、selection 已失效仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-034 | e2e | P2 | `src/officecli/CommandBuilder.Mark.cs` | watch mark/list/get-marks/clear | 真实 watch 下 explicit path mark、marks JSON、`unmark --path` 清理已固定；`selected` 多目标、`--all` 和 stale mark 仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-035 | e2e | P2 | `src/officecli/CommandBuilder.Goto.cs` | watch `goto` scroll target | 真实 watch 下 paragraph、table、row、cell anchor 成功，缺失 paragraph 返回错误已固定；无 watch 分支仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-036 | e2e | P1 | `src/officecli/CommandBuilder.Refresh.cs` | `refresh` 降级协议 | 非 docx 返回 `unsupported_type`；当前平台无 Word/HTML fallback 时返回 `refresh_failed`，有 fallback 时返回 `Refreshed` + backend envelope；字段 fixture 的 backend/Pages 后置状态已固定；真实字段结果差异仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-037 | e2e | P1 | `src/officecli/CommandBuilder.cs` | `merge` 模板替换 | body/table/header/footer 中占位符替换、缺失 key 保留并报告、数组/对象/literal key flatten、输出存在无 `--force` 报错、非法 JSON 报错、notes/comments part 替换和跨 run 占位符当前不替换均已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-038 | e2e | P1 | `src/officecli/CommandBuilder.Plugins.cs` | plugin lint/dump-reader 对 docx schema | 合法 dump-reader paragraph prop 的 lint success/`unknown_prop_count=0`，以及 JSON array 的 `corrupt_batch` 已固定；unknown prop、非 dump-reader plugin 和 Windows host 仍待补；见 `tests/OfficeCli.Tests/Integration/WordPluginContractTests.cs` | 部分自动覆盖 |
| WORD-CMD-039 | e2e | P0 | `.github/workflows/build.yml`, `src/officecli/officecli.csproj` | 发布后二进制 Word smoke | macOS `osx-arm64` 和 Windows `win-x64` self-contained/single-file 产物生成、非空已固定；当前可执行平台进一步跑 `create/add/view text/validate`，并关闭自动 resident/update；其他架构、签名/installer/notarization 仍由 CI/发布流程覆盖；见 `tests/OfficeCli.Tests/E2E/WordPublishedBinarySmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-040 | e2e | P1 | `src/officecli/CommandBuilder.IntegrationStubs.cs`, MCP server | MCP/agent Word 调用面 | 真实 MCP stdio `initialize/tools/list/tools/call`、未知工具错误，以及通过 `officecli` 工具执行 Word `create/add/query` 的 JSON envelope 已固定；batch/dump 与外部 agent SDK 适配仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-041 | e2e | P1 | `src/officecli/CommandBuilder.Add.cs` | `swap` CLI 面 | 两个 body 段落、同一表格行内 cell、非法跨父级 `invalid_value`、不存在 path `not_found`，以及 resident forward 后内存/close 后磁盘顺序均已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-042 | e2e | P1 | `src/officecli/CommandBuilder.View.cs` | view 高成本/插件分支 | `view forms` 的 protection/fields/sdt JSON 结构和未知 mode 的 `invalid_value` envelope 已固定；`view pdf` exporter、`--page-count` 降级和 resident-lock 处理仍待补；screenshot 的 grid/native 基础错误边界见 WORD-CMD-047；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-043 | e2e | P1 | `src/officecli/CommandBuilder.cs` | `create` 的 `--type/--locale/--minimal` 选项矩阵 | 无扩展名时 `--type=docx` 补扩展名；locale 写入对应默认字体并对 RTL locale 设置布局；minimal 与标准 docx 的 baseline parts 差异稳定；四种产物均可 reopen/validate；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-044 | e2e | P1 | `src/officecli/CommandBuilder.GetQuery.cs` | `get` 默认 path、`--depth`、深度上限和 error node | `/`、`/body` 的默认 depth 与显式 depth 结构稳定；超大 depth 被安全限制；不存在节点统一为非零退出和当前 JSON message/error 形状，文件不被修改；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-045 | e2e | P1 | `src/officecli/CommandBuilder.View.cs` | `view text/annotated` 截断与 `stats/issues` 过滤 | `--start/--end/--max-lines` 的边界、总数/截断标记、annotated 列稳定；`issues --type/--limit` 过滤和限制数量稳定；JSON stdout 不混入诊断日志；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs`、`tests/OfficeCli.Tests/Functional/WordFieldViewContractTests.cs` | 已有自动覆盖 |
| WORD-CMD-046 | e2e | P1 | `src/officecli/CommandBuilder.View.cs`, `WordHandler.HtmlPreview.*.cs` | `view html --page --out` | 默认 stdout 与 `--out` 文件内容等价；HTML 保留 paragraph `data-path`、复杂对象 anchor 和非 ASCII 文本；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-CMD-047 | e2e | P2 | `src/officecli/CommandBuilder.View.cs` | `view screenshot/pdf` 输出与降级 | screenshot 指定页/尺寸、HTML `--grid`、当前宿主 native 不可用错误和 pdf exporter 成功/失败输出已固定；Windows + Word native、resident-lock 组合仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-048 | integration | P1 | `src/officecli/CommandBuilder.Batch.cs` | batch 文件、stdin、JSONL 和坏输入 | `--commands`、`--input <file>`、stdin、`--input -` 四种输入的成功结果等价；空数组、UTF-8 BOM、JSONL 多行当前 `invalid_json`、`--commands` 与 `--input` 同传当前 `internal_error`、顶层 object/未知字段边界均已固定；失败不污染目标 docx；stdin 重定向 warning 仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |

## 文档结构与文本格式

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-DOC-001 | contract | P0 | `schemas/help/docx/document.json` | 设置文档属性 `author/title/keywords/description/lastModifiedBy` | `get /` 返回 canonical key；别名 `creator` 可写入 author | 已有自动覆盖 |
| WORD-DOC-002 | contract | P1 | `schemas/help/docx/document.json` | 设置 `docDefaults.*`、CJK grid、compatibility | `docDefaults` effective 值、`docGrid.*`、`charSpacingControl`、`compatibility.mode`、compatibility flag/preset 已固定；单元见 `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs` | 已有单元覆盖 |
| WORD-DOC-003 | integration | P0 | `schemas/help/docx/body.json` | body 下段落、表格、section 顺序读取 | `get /body --depth 1` 返回有序子节点 | 已有自动覆盖 |
| WORD-DOC-004 | contract | P0 | `schemas/help/docx/paragraph.json` | 新增段落文本、样式、对齐、间距、缩进 | `get` 返回 `align/spaceBefore/spaceAfter/lineSpacing` 等 canonical key | 已有自动覆盖 |
| WORD-DOC-005 | contract | P1 | `schemas/help/docx/paragraph.json` | 段落有效样式继承 `effective.*` | 直接属性缺省时出现 `effective.X` 和 `effective.X.src` | 已有自动覆盖 |
| WORD-DOC-006 | integration | P1 | `examples/word/paragraph-formatting.*` | 段落格式示例转集成测试 | 示例生成成功；关键段落属性可通过 `get` 验证 | 示例待改造 |
| WORD-DOC-007 | contract | P0 | `schemas/help/docx/run.json` | run 文本、字体、字号、粗斜体、颜色、下划线 | `get` 返回格式值；`view text` 保持文字顺序 | 已有自动覆盖 |
| WORD-DOC-008 | contract | P1 | `schemas/help/docx/run.json` | 复杂脚本 / 东亚字体 / RTL run 属性 | `get` 返回对应 lang/font slot、size.cs、bold.cs、italic.cs、direction=rtl；RTL 不影响普通文本读取 | 已有自动覆盖 |
| WORD-DOC-009 | integration | P1 | `examples/word/run-formatting.*` | run 格式示例转集成测试 | 示例生成成功；关键 run 属性可读 | 示例待改造 |
| WORD-DOC-010 | contract | P1 | `schemas/help/docx/pagebreak.json` | 添加 page/column break | `query pagebreak` 返回类型；`view outline/text` 不丢正文 | 已有自动覆盖 |
| WORD-DOC-011 | contract | P1 | `schemas/help/docx/tab.json` | 段落 tab stop 添加、修改、删除 | `get` 返回位置、align、leader；删除后读取模型不再出现，但当前实现会留下空 `w:tabs` schema 错误 | 已有自动覆盖 |
| WORD-DOC-012 | contract | P2 | `schemas/help/docx/ptab.json` | header/footer 中 positional tab | body 段落 ptab add/set/get 已固定；header/footer 场景待补 | 部分自动覆盖 |
| WORD-DOC-013 | contract | P0 | `schemas/help/docx/style.json` | 新增 paragraph/character/table style | `get /styles/<id>` 返回类型、name、basedOn、格式属性；单元见 `tests/OfficeCli.Tests/Unit/WordStyleAndNumberingTests.cs` | 已有单元覆盖 |
| WORD-DOC-014 | contract | P1 | `schemas/help/docx/styles.json` | styles 容器查询和添加 style | `query style` 包含新增样式；重复 custom id 抛错且不污染 styles | 已有自动覆盖 |
| WORD-DOC-015 | integration | P1 | `examples/word/document-formatting.*` | 页面背景、默认字体、metadata 等文档格式示例 | 关键 document props 可读；`validate` 通过 | 示例待改造 |
| WORD-DOC-016 | contract | P0 | `schemas/help/docx/section.json` | section 尺寸、边距、方向、栏、页码、RTL gutter | `get /body/sectPr[1]` 返回 canonical 长度和方向值 | 已有自动覆盖 |
| WORD-DOC-017 | integration | P1 | `examples/word/sections.*` | 多 section 示例转集成测试 | section 数量、orientation、margin、columns 可读 | 示例待改造 |
| WORD-DOC-018 | contract | P1 | `schemas/help/docx/header.json` | 添加 default/first/even header | duplicate type 被拒绝；section 引用正确；内容可 query | 已有自动覆盖 |
| WORD-DOC-019 | contract | P1 | `schemas/help/docx/footer.json` | 添加 default/first/even footer | 与 header 语义一致；页脚内容可读写 | 已有自动覆盖 |
| WORD-DOC-020 | contract | P1 | `schemas/help/docx/run.json` | run 高级格式 contract | highlight、strike、caps、baseline position、kern、charspacing、text effects、theme color 的 add/set/get/readback 与 schema 一致；高级开关、w14、shading/theme 已有 unit，contract/schema parity 待补 | 部分自动覆盖 |
| WORD-DOC-021 | contract | P1 | `schemas/help/docx/paragraph.json`, `schemas/help/docx/run.json`, `schemas/help/docx/revision.json` | range/find 格式化与 tracked find contract | `find` 文本匹配与 regex replace、`range` 半开区间和跨段落 `/body` scope、`find+range`/`text+range`/无格式 range 错误，以及 `find+revision.author` 的 replacement（del+ins）、空 replacement（del-only）、format-only marker 已固定；见 `tests/OfficeCli.Tests/Functional/WordFindReplaceContractTests.cs` | 已有自动覆盖 |
| WORD-DOC-022 | regression | P0 | `schemas/help/docx/*.json`, `officecli help docx` | schema/help/runtime 漂移 | `revision` schema 契约、`help docx set revision` 路由、`help docx all` 汇总，以及全部嵌入 docx schema 元素与五位 operations/path 摘要一致已固定；其余元素的全量 property/runtime parity 待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-DOC-023 | integration | P1 | `examples/word/*.sh`, `examples/word/*.py` | Word 示例脚本 smoke | `content-controls.sh`、`pictures.sh`、`revisions.sh`、`tables.sh` 四个代表 shell 示例已在临时目录执行，生成 docx 可重新打开并查询 paragraph；其余 shell/Python SDK 示例和每个示例的专属关键节点断言仍待补；见 `tests/OfficeCli.Tests/Integration/WordExampleScriptSmokeTests.cs` | 部分自动覆盖 |
| WORD-DOC-024 | integration | P1 | `src/officecli/CommandBuilder.cs`, OOXML package parts | 已有 `.docx` 的未知 part 保留 | customXml、customXmlProperties、fontTable、webSettings 在 no-op/local mutation 和 `dump -> batch` 后的内容、关系和 validate 已固定；主题、嵌入对象、未知 docProps 的完整 bytes/content-types 矩阵仍待补；见 `tests/OfficeCli.Tests/Integration/WordDocumentFidelityTests.cs` | 部分自动覆盖 |
| WORD-DOC-025 | contract | P1 | `schemas/help/docx/header.json`, `footer.json`, `section.json` | 多 section 与 header/footer 关联 | 多 section 的 default/first/even header/footer reference、part 类型、titlePg/evenAndOdd toggle、section 归属和 round-trip 已固定；link-to-previous 开关和 section 增删矩阵仍待补；见 `tests/OfficeCli.Tests/Functional/WordSectionRelationshipTests.cs` | 部分自动覆盖 |
| WORD-DOC-026 | contract | P1 | `schemas/help/docx/field.json`, `toc.json`, `fieldchar.json` | field cached result、dirty 状态和刷新前后 readback | `field_not_evaluated`/`field_cache_stale`、`view text` sentinel、PAGE/NUMPAGES dirty readback，以及 refresh backend/Pages 状态已固定；TOC/PAGE/SEQ/REF 实际结果差异、updateFields 和 Word backend 矩阵仍待补；见 `tests/OfficeCli.Tests/Functional/WordFieldViewContractTests.cs`、`tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |

## 表格、列表与编号

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-TBL-001 | contract | P0 | `schemas/help/docx/table.json` | 新增 table `rows/cols/border/layout/colWidths` | `query table` 返回目标；行列数、边框、布局可读 | 已有自动覆盖 |
| WORD-TBL-002 | contract | P0 | `schemas/help/docx/table-row.json` | 添加、设置、删除 row | 行数变化正确；`header/height` 可读 | 已有自动覆盖 |
| WORD-TBL-003 | contract | P0 | `schemas/help/docx/table-cell.json` | cell 文本、宽度、shd、align、valign、padding | `get` 返回 cell 与内部段落/run 格式；文本不丢失 | 已有自动覆盖 |
| WORD-TBL-004 | contract | P1 | `schemas/help/docx/table-cell.json` | `vmerge/gridspan/hmerge` 合并单元格 | 合并后 cell 索引变化符合说明；被吸收 cell 不可再 get | 已有自动覆盖 |
| WORD-TBL-005 | contract | P1 | `schemas/help/docx/table-column.json` | 虚拟 column add/remove | gridCol 与每行 cell 同步；列宽可读 | 已有自动覆盖 |
| WORD-TBL-006 | integration | P1 | `examples/word/tables.*` | 表格示例转集成测试 | 7 个 Word 表格生成；合并、热力色、RTL、边框关键属性可读 | 示例待改造 |
| WORD-TBL-007 | contract | P0 | `schemas/help/docx/numbering.json` | numbering 容器读取和查询 | `get /numbering` 返回 abstractNum/num 子节点 | 已有自动覆盖 |
| WORD-TBL-008 | contract | P0 | `schemas/help/docx/abstractNum.json` | 新增 numbering 模板 | 9 级默认 level 存在；样式和格式可读 | 已有自动覆盖 |
| WORD-TBL-009 | contract | P0 | `schemas/help/docx/num.json` | 新增 numbering instance | `abstractNumId` 引用正确；重复/非法引用有报错 | 已有自动覆盖 |
| WORD-TBL-010 | contract | P1 | `schemas/help/docx/level.json` | 设置 level 格式、起始编号、缩进、文本模板 | `get` 返回 `format/start/text/indent` 等关键值 | 已有自动覆盖 |
| WORD-TBL-011 | integration | P1 | `examples/word/numbering.*` | 列表和多级编号示例转集成测试 | ordered/bullet/readback 正确；中断后重启符合说明 | 示例待改造 |

## 媒体、绘图与复杂对象

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-MEDIA-001 | integration | P1 | `schemas/help/docx/picture.json` | CLI 查询 inline picture | handler 和真实 CLI `query picture` 均返回 `wrap=inline`；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-MEDIA-002 | integration | P1 | `schemas/help/docx/picture.json` | CLI 查询 floating picture `wrap/hPosition/vPosition` | handler 和真实 CLI `query picture` 均返回 wrap、h/v position 和 relative frame；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-MEDIA-003 | contract | P0 | `schemas/help/docx/picture.json` | add picture `src/width/height/alt` | `get` 返回尺寸和 alt；`validate` 通过 | 已有自动覆盖 |
| WORD-MEDIA-004 | contract | P1 | `schemas/help/docx/picture.json` | crop、decorative、link、behindText、wrap variants | wrap variants、floating position、crop、decorative、link、behindText 写入固定 | 已有自动覆盖 |
| WORD-MEDIA-005 | integration | P1 | `examples/word/pictures.*` | 图片示例转集成测试 | 8 个场景生成；关键图片路径和属性可读 | 示例待改造 |
| WORD-MEDIA-006 | integration | P1 | `schemas/help/docx/ole.json` | CLI 查询 OLE 对象并与 picture 区分 | handler 和真实 CLI `query ole` 均返回 `type=ole`、`progId`，不会混入 picture；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-MEDIA-007 | contract | P1 | `schemas/help/docx/ole.json` | add OLE / embedded package | OLE rel、contentType、fileSize、尺寸可读；dump-batch 后 rel 不悬空待补 | 部分自动覆盖 |
| WORD-MEDIA-008 | contract | P1 | `schemas/help/docx/chart.json` | 添加 chart 和基础属性 | `query chart` 返回类型、标题、数据范围；`validate` 通过；单元见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-MEDIA-009 | contract | P1 | `schemas/help/docx/chart-series.json` | 添加/修改 chart series | series 名称和值修改可读；range-backed categories/values 和 series color 读回固定；单元见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-MEDIA-010 | contract | P2 | `schemas/help/docx/chart-axis.json` | chart axis set/get | value axis title、min/max、number format 可读；单元见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-MEDIA-011 | integration | P1 | `examples/word/charts.*` | 图表示例转集成测试 | 图表数量、series、axis 关键属性可读 | 示例待改造 |
| WORD-MEDIA-012 | contract | P1 | `schemas/help/docx/equation.json` | LaTeX-ish equation add/get/set/remove | inline/display 模式、formula readback、set/remove 当前行为固定；单元见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 已有单元覆盖 |
| WORD-MEDIA-013 | integration | P1 | `examples/word/formulas.*` | 公式示例转集成测试 | 公式数量和关键公式文本可读 | 示例待改造 |
| WORD-MEDIA-014 | contract | P1 | `schemas/help/docx/diagram.json` | mermaid diagram add | native diagram 返回 `/body/group[N]`；当前 `/body/textbox[N]` 仅暴露 grouped drawing 的首个文本框；group resize/remove 当前行为固定；`validate` 通过；单元见 `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | 已有单元覆盖 |
| WORD-MEDIA-015 | integration | P2 | `examples/word/diagram.*` | diagram 示例转集成测试 | 关键图形或渲染输出存在；复杂视觉先不做像素断言 | 示例待改造 |
| WORD-MEDIA-016 | contract | P1 | `schemas/help/docx/textbox.json` | textbox add/set/get | 内容树可通过 `/body/textbox[N]/p/r` 寻址；`fill/line/width/height/geometry` set 面可执行；`text/position` add-only 边界固定；单元见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 已有单元覆盖 |
| WORD-MEDIA-017 | integration | P1 | `examples/word/textbox.*` | 文本框示例转集成测试 | textbox 数量和关键属性可读 | 示例待改造 |
| WORD-MEDIA-018 | contract | P2 | `schemas/help/docx/shape.json` | floating shape add/set/get/remove | raw drawing tree、geometry/fill/size 子路径读回；`fill/line/width/height/geometry` set 面可执行；position add-only 边界和 remove 后不再可导航固定；单元见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 已有单元覆盖 |
| WORD-MEDIA-019 | contract | P1 | `schemas/help/docx/watermark.json` | text/image watermark | text watermark add/get/query/set/remove、VML 属性读回、非法 rotation 固定；handler 层 `image` 当前退回默认 text watermark 的行为已固定，真实 image watermark 待补；单元见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 部分自动覆盖 |
| WORD-MEDIA-020 | regression | P0 | `README.md` | `dump -> batch` 保留图片 part 和 rel | 新文档图片可 `query/get`、image part bytes 与源 data URI 一致，且 `validate` 通过 | 已有自动覆盖 |
| WORD-MEDIA-021 | regression | P0 | `README.md` | `dump -> batch` 保留 OLE / embedded package | `query ole` 返回 progId/name/contentType/fileSize；`validate` 通过 | 已有自动覆盖 |
| WORD-MEDIA-022 | regression | P1 | `README.md` | `dump -> batch` 保留 chart 的 externalData / embedded workbook | chart/series/cache 可读且 `validate` 通过；workbook rel 存在待补 | 部分自动覆盖 |
| WORD-MEDIA-023 | regression | P1 | `README.md` | header/footer 内图片或 textbox round-trip | header/footer 内图片 dump/batch 后可在各自子树读回 name/alt 已固定；textbox 待补；见 `tests/OfficeCli.Tests/Integration/WordDumpBatchRoundTripTests.cs` | 部分自动覆盖 |
| WORD-MEDIA-024 | regression | P1 | `src/officecli/CommandBuilder.GetQuery.cs` | picture/OLE `get --save` round-trip | handler 级 payload bytes/contentType、CLI picture/OLE JSON 的 contentType/savedBytes/savedTo、无 payload exit/envelope 已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 已有自动覆盖 |
| WORD-MEDIA-025 | regression | P1 | `src/officecli/Handlers/Word/WordBatchEmitter.Resources.cs` | dump 资源完整性 | full dump 的 numbering/styles/theme/settings/core/app docProps 在 body 前且无 warning 已固定；fontTable/webSettings/customXml/custom docProps、资源 replay 后关键 raw 保留仍待补；见 `tests/OfficeCli.Tests/Unit/WordDumpEmitterTests.cs` | 部分自动覆盖 |
| WORD-MEDIA-026 | visual | P2 | `src/officecli/Handlers/Word/WordHandler.HtmlPreview.*.cs` | HTML preview 复杂对象 | chart/shape/textbox/watermark 的 HTML SVG、文本、watermark class 和 paragraph `data-path` 已固定；revision marker、comment/note marker 的稳定 anchor/class 仍待补；见 `tests/OfficeCli.Tests/Functional/WordHtmlPreviewContractTests.cs` | 部分自动覆盖 |
| WORD-MEDIA-027 | regression | P1 | `src/officecli/Handlers/WordHandler.cs`, `src/officecli/Handlers/Word/WordBatchEmitter.Resources.cs` | ActiveX/VML/复杂 carrier round-trip | native chart carrier 的 host relId 重写、chart style child part、external relationship 和 chart XML `dump→batch` 已固定；ActiveX form control、VML textbox/shape、rich SDT、equation OLE、diagram relIds、chart userShapes、embedded fonts、numPicBullet 仍待补；见 `tests/OfficeCli.Tests/Integration/WordComplexDumpCarrierRoundTripTests.cs` | 部分自动覆盖 |

## 引用、字段、批注与表单

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-REF-001 | contract | P0 | `schemas/help/docx/hyperlink.json` | 外部 URL hyperlink add/set/get/remove | rel 存在；`get` 返回 url/text/style | 已有自动覆盖 |
| WORD-REF-002 | contract | P1 | `schemas/help/docx/hyperlink.json` | 内部 bookmark anchor hyperlink | anchor 指向合法 bookmark；非法目标有报错 | 已有自动覆盖 |
| WORD-REF-003 | contract | P1 | `schemas/help/docx/bookmark.json` | bookmark start/end 成对创建和改名 | `query bookmark` 返回 name；非法 name 被拒绝 | 已有自动覆盖 |
| WORD-REF-004 | contract | P1 | `schemas/help/docx/toc.json` | TOC field 插入和属性 | `query toc` 返回 field；`refresh` 可选更新 | 已有自动覆盖 |
| WORD-REF-005 | contract | P1 | `schemas/help/docx/field.json` | complex field add/get/set/remove | begin/instr/separate/result/end 结构完整；`query field` 返回的 `/field[N]` 当前为虚拟路径，直接 remove 被拒绝；见 `tests/OfficeCli.Tests/Unit/WordFieldAndFormTests.cs` | 已有单元覆盖 |
| WORD-REF-006 | contract | P1 | `schemas/help/docx/instrtext.json` | instrText query/set/remove | 指令文本可读写；remove `instrText` 实际 run 后 instruction 和折叠 field 查询消失 | 已有单元覆盖 |
| WORD-REF-007 | contract | P2 | `schemas/help/docx/fieldchar.json` | fieldChar get/query/remove | begin/separate/end 类型可读；remove 单个 marker run 后目标 marker 消失，折叠 field 查询消失 | 已有单元覆盖 |
| WORD-REF-008 | integration | P1 | `examples/word/fields.*` | fields 示例转集成测试 | PAGE/REF/SEQ 等关键 field 可查询 | 示例待改造 |
| WORD-REF-009 | contract | P1 | `schemas/help/docx/footnote.json` | footnote add/set/get/remove | 正文引用和 FootnotesPart note 成对存在 | 已有自动覆盖 |
| WORD-REF-010 | contract | P1 | `schemas/help/docx/endnote.json` | endnote add/set/get/remove | 正文引用和 EndnotesPart note 成对存在 | 已有自动覆盖 |
| WORD-REF-011 | contract | P1 | `schemas/help/docx/comment.json` | comment add/set/get/query/remove | comment range 与 CommentsPart 内容一致；author/date 可读 | 已有自动覆盖 |
| WORD-REF-012 | contract | P1 | `schemas/help/docx/formfield.json` | text/check/dropdown form field | bookmark namespace 正确；默认值和状态可读 | 已有自动覆盖 |
| WORD-REF-013 | contract | P0 | `schemas/help/docx/sdt.json` | SDT schema type values 与 add-time 实现一致 | 未实现 type 不出现在 schema；`type` 不可 set | 已有单元覆盖 |
| WORD-REF-014 | contract | P0 | `schemas/help/docx/sdt.json` | SDT canonical `type` add/get | `add --prop type=dropdown` 成功；`get` 返回 `type` 不返回 `sdtType` | 已有单元覆盖 |
| WORD-REF-015 | contract | P0 | `schemas/help/docx/sdt.json` | SDT lock -> editable readback | `editable` 随 `lock` 正确变化，重开文件后保持 | 已有单元覆盖 |
| WORD-REF-016 | contract | P0 | `schemas/help/docx/sdt.json` | checkbox SDT checked state | `type=checkbox`；`checked=true/false` 可读 | 已有单元覆盖 |
| WORD-REF-017 | contract | P1 | `schemas/help/docx/sdt.json` | dropdown/combobox/date/picture/group/richtext | 每类 add 成功；类型特定 props、placeholder、stable sdtId 路径可读 | 已有自动覆盖 |
| WORD-REF-018 | integration | P1 | `examples/word/content-controls.*` | content controls 示例转集成测试 | 8 类 SDT 均存在；关键 props 可读 | 示例待改造 |
| WORD-REF-019 | contract | P2 | `schemas/help/docx/permStart.json` | editing permission range marker | permStart/permEnd 成对；id、edGrp、ed、colFirst/colLast 可读；当前返回 `[@id]` 路径但读删使用 positional 路径；非法 id 固定 | 已有自动覆盖 |
| WORD-REF-020 | contract | P1 | `schemas/help/docx/hyperlink.json` | hyperlink 完整格式属性 | tooltip、tgtFrame、history=false、docLocation、rStyle、color/theme、font/font.cs、size、bold/italic、underline/underline.color、strike、highlight readback 当前行为已固定；visited 独立语义当前未暴露 | 已有自动覆盖 |
| WORD-REF-021 | contract | P1 | `schemas/help/docx/comment.json`, `src/officecli/Handlers/Word/WordHandler.CommentsExt.cs` | comment 扩展属性 | `author/initials/date/id/done/parentId/direction/anchoredTo` 的 handler readback、`done/resolved` alias、reply thread parentId 映射、`done/parentId` selector 查询、rangeOpen/rangeEnd 跨段 anchor、remove 后 orphan range marker 清理已固定；真实 CLI 当前不暴露 `direction`，`comment[direction=rtl]` 返回 `unknown_key`/空结果也已固定；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` 与 `tests/OfficeCli.Tests/Functional/WordNotesAndCommentsContractTests.cs` | 部分自动覆盖 |
| WORD-REF-022 | contract | P1 | `schemas/help/docx/footnote.json`, `schemas/help/docx/endnote.json` | note 格式属性 | footnote/endnote text、id、`font/size/bold/italic/color/underline/strike/highlight` add/get readback、direction 的 Format readback、align 的 raw XML 写入、remove 后 query 消失且 validate 空结果已固定；align Format readback、set 格式、reference mark dump 载体和正文 reference run 格式待补；见 `tests/OfficeCli.Tests/Functional/WordNotesAndCommentsContractTests.cs` | 部分自动覆盖 |
| WORD-REF-023 | contract | P1 | `schemas/help/docx/formfield.json` | formfield 更新面 | text/check/dropdown set 后状态、默认值、选项列表和 bookmark wrapper 当前 readback 固定；dropdown 非法 text 当前只改缓存显示文本、不改 `result`，`result` 作为 unsupported property 返回且不改选择索引；见 `tests/OfficeCli.Tests/Functional/WordFormFieldContractTests.cs` | 已有自动覆盖 |
| WORD-REF-024 | e2e | P2 | `schemas/help/docx/toc.json`, `field.json`, `src/officecli/CommandBuilder.Refresh.cs` | 真实 refresh 的 TOC/PAGE/SEQ/REF 行为 | field fixture 的 refresh backend、成功消息、扩展属性页数和不可用 backend 的降级已固定；TOC/PAGE/SEQ/REF 各字段 result/raw XML 的前后差异仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |

## 修订与回归场景

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-REV-001 | contract | P0 | `schemas/help/docx/revision.json` | run insertion/deletion/format revision | `query revision` 返回 type、author、id、path、text，显式 date 归一化 readback 已固定 | 已有自动覆盖 |
| WORD-REV-002 | contract | P0 | `schemas/help/docx/revision.json` | paragraph add/remove/set with `revision.author` | paragraph format revision 可查询；synthetic revision 当前读回 `revision.type=paragraph`，host paragraph 读回 `format`；当前 paragraph `revision.type=ins` via set 明确拒绝；remove 的 paragraph-mark/run deletion 和 run accept/reject 已固定；tracked paragraph add 仍明确不支持；见 `tests/OfficeCli.Tests/Unit/WordRevisionTests.cs`、`tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | 部分自动覆盖 |
| WORD-REV-003 | contract | P1 | `schemas/help/docx/revision.json` | moveFrom/moveTo 配对 `revision.id` | run 级 moveFrom/moveTo 两半共享 id；synthetic revision path 按 type 消歧；单元已固定；range marker 细节和更多 pair 冲突场景待补 | 部分自动覆盖 |
| WORD-REV-004 | contract | P1 | `schemas/help/docx/revision.json` | table/row/cell scope revision | table/cell format marker、row 级 ins marker 可查询，synthetic row type 当前为 `rowIns`、host row 为 `ins`；row/cell tracked remove marker 已固定；trPr format 和 cell accept/reject 待补；见 `tests/OfficeCli.Tests/Unit/WordRevisionTests.cs`、`tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | 部分自动覆盖 |
| WORD-REV-005 | contract | P1 | `schemas/help/docx/revision.json` | section property revision | `sectPrChange` author/id 可读；section orientation 更新可读；synthetic revision type 当前为 `format` | 已有自动覆盖 |
| WORD-REV-006 | contract | P1 | `schemas/help/docx/revision.json` | accept/reject 单个 revision | run insertion/deletion/format/move accept/reject、nativePath 动作、按 author/type 批量筛选、stale id 和 action/creation key 混用错误当前结果固定；见 `tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | 已有自动覆盖 |
| WORD-REV-007 | integration | P1 | `examples/word/revisions.*` | revisions 示例转集成测试 | 8 个 revision 场景生成；accept/reject temp copy 验证 | 示例待改造 |
| WORD-REV-008 | regression | P0 | `tests/OfficeCli.Tests/UnitTest1.cs` | picture 与 OLE query 不互相污染 | picture 数量和 ole 数量稳定 | 已有单元覆盖 |
| WORD-REV-009 | regression | P0 | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | SDT schema/runtime 漂移防护 | schema 声明和 handler readback 不漂移 | 已有单元覆盖 |
| WORD-REV-010 | regression | P1 | `README.md` | schema 中 canonical key 与 legacy alias | shape set 的 `preset/fillcolor/linecolor/linewidth` legacy alias 写入后 raw canonical 输出稳定；其他 alias 待补 | 部分自动覆盖 |
| WORD-REV-011 | regression | P1 | `README.md` | 错误路径、非法属性、非法枚举值 | chart series 非法数值抛 `invalid_value` 但当前会清空 values 并留下 schema 错误；其他场景待补 | 部分自动覆盖 |
| WORD-REV-012 | contract | P1 | `schemas/help/docx/revision.json`, `src/officecli/Handlers/Word/WordHandler.Mutations.cs` | tracked remove paragraph/run/table 当前行为 | `remove` 携带 `revision.author/id/date` 时 run 的 `w:del`、paragraph-mark/run deletion、row/cell structural deletion marker、host/revision readback 和 run accept/reject 已固定；其他 table scope 的 accept/reject 组合待补；见 `tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | 部分自动覆盖 |
| WORD-REV-013 | contract | P1 | `schemas/help/docx/revision.json` | revision date/id 冲突 | 显式 date 归一化、缺 date 自动生成、重复 id 通过 type-specific path 消歧已固定；nativePath/selector action 的 stale id 与混用保护已固定；move pair id 冲突仍待补；见 `tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | 部分自动覆盖 |
| WORD-REV-014 | regression | P1 | `schemas/help/docx/*.json` | alias/canonical 全量回放 | document/paragraph/run/hyperlink/table-cell/field 六类高频 alias 写入后 canonical readback 已固定；picture 的 `w/h` 当前未归一化且会落入错误属性解析，其余 schema alias 的全量回放、legacy key 不泄漏规则仍待补；见 `tests/OfficeCli.Tests/Functional/WordAliasContractTests.cs` | 部分自动覆盖 |
| WORD-REV-015 | contract | P1 | `schemas/help/docx/revision.json`, `src/officecli/Handlers/Word/WordHandler.Helpers.FindReplace.cs` | find + revision tracked replacement | `revision.author` 驱动每个匹配的 del/ins、空 replacement 只产生 del、format-only 产生 `rPrChange`；author/date/id、text 和 validate 当前行为固定；见 `tests/OfficeCli.Tests/Functional/WordFindReplaceContractTests.cs` | 已有自动覆盖 |

## 负向与边界测试

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-NEG-001 | contract | P0 | `schemas/help/docx/*.json` | 非法 path：不存在节点、越界索引、错误父节点 | 返回 `not_found` 或等价结构化错误；文件不变 | 已有单元覆盖 |
| WORD-NEG-002 | contract | P0 | `schemas/help/docx/*.json` | 非法 enum / bool / length / color 值 | 返回 `invalid_value`；错误信息包含合法范围或示例 | 已有自动覆盖 |
| WORD-NEG-003 | contract | P0 | `schemas/help/docx/*.json` | 缺少 required prop，例如 picture `src`、sdt `type` | 命令失败；不创建半成品节点或 dangling rel | 已有自动覆盖 |
| WORD-NEG-004 | contract | P1 | `schemas/help/docx/*.json` | 不支持属性或只读属性 set | paragraph unsupported 属性返回 unsupported 且不应用该属性；当前实现可能刷新 textId，不承诺 XML 完全不变 | 已有自动覆盖 |
| WORD-NEG-005 | contract | P1 | `schemas/help/docx/bookmark.json` | 重复 bookmark/formfield 名称或非法名称 | bookmark 重名保留；formfield 非法名被拒绝且不新增 formfield | 已有自动覆盖 |
| WORD-NEG-006 | contract | P1 | `schemas/help/docx/header.json` | 重复添加同一 section 的同类型 header/footer | default/first/even header/footer 重复添加均抛错，原有 part 不丢失；见 `tests/OfficeCli.Tests/Unit/WordHeaderFooterTests.cs` | 已有单元覆盖 |
| WORD-NEG-007 | integration | P1 | `README.md` | batch 中间失败：默认继续与 `--stop-on-error` | 默认后续命令继续；`--stop-on-error` 停在失败项；结果可解释 | 已有自动覆盖 |
| WORD-NEG-008 | regression | P1 | `README.md` | `raw-set`、`dump-batch` 失败时避免关系悬空 | raw-set 失败后不新增 validate 错误且已存在 chart relationship 仍可解析；dump-batch 失败路径待补 | 部分自动覆盖 |
| WORD-NEG-009 | integration | P2 | `README.md` | 打开损坏或非 docx 文件 | 损坏 `.docx` 打开失败、CLI JSON 返回 `corrupt_file`，且原始 bytes 不被覆盖 | 已有自动覆盖 |
| WORD-NEG-010 | integration | P0 | `src/officecli/CommandBuilder.cs` | docx protection gate | 受保护文档的 add/set 默认拒绝且 `--force` 绕过；batch 内 add/remove/raw-set 默认拒绝且 `--force` 绕过；standalone remove/raw-set 当前不走 protection gate 并会执行；允许编辑区域内 mutation 规则待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-NEG-011 | e2e | P1 | `src/officecli/CommandBuilder*.cs` | CLI 文件/路径错误 envelope | 缺文件 `file_not_found`、扩展名不支持 `unsupported_type`、路径越界 `not_found`、无 watch 的 `get selected` 当前 `success=false` + `message` 且无 `error`、bare selector `bare_selector_rejected` 已固定；resident pipe busy 待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-NEG-012 | integration | P1 | `src/officecli/CommandBuilder.Add.cs`, `src/officecli/CommandBuilder.Set.cs` | prop 自动纠错和 warning | bare `key=value` 缺 `--prop` 时 exit 2、stderr warning、mutation 继续但属性不应用，大小写 key 和 `alignment` legacy alias 的 canonical readback、unsupported list/set warning 已固定；拼写自动纠错仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-NEG-013 | regression | P1 | `src/officecli/Handlers/Word/WordHandler.Mutations.cs` | 删除/移动复杂关系不悬空 | 删除 picture/OLE/chart/hyperlink/comment/footnote/endnote/header/footer 后 query 清空且 rel、part、range marker 不留下 validate 错误已固定；复杂 carrier 删除仍待补；见 `tests/OfficeCli.Tests/Integration/WordComplexRelationshipCleanupTests.cs` | 部分自动覆盖 |
| WORD-NEG-014 | regression | P1 | `src/officecli/Handlers/Word/WordHandler.Mutations.cs`, `CommandBuilder.Batch.cs` | mutation 失败原子性 | failed set、raw-set 非法 XML、in-process malformed batch 的 bytes/readback/validate 不变已固定；resident move/copy 的 flush failure 边界已固定；remove/move/swap、关系冲突和 CLI batch 全量矩阵仍待补；见 `tests/OfficeCli.Tests/Integration/WordMutationAtomicityTests.cs`, `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-NEG-015 | e2e | P1 | `src/officecli/CommandBuilder.cs`, `ResidentClient`, `ResidentServer` | resident 锁竞争、busy、stale marker 清理 | resident 占用时 `create --force` 返回 `file_locked`；resident 异常退出后 pipe 不再可达且下一次 create 可恢复；merge/pdf/raw busy 和残留 marker 矩阵仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-NEG-016 | e2e | P2 | `src/officecli/CommandBuilder*.cs` | Unicode、空白、长路径和平台路径边界 | CJK/RTL 文件名、空格路径、CJK/RTL 文本在当前宿主 create/add/query/view/validate 中已固定；长路径、Windows 风格分隔符和 macOS/Windows 发布包矩阵仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-NEG-017 | e2e | P1 | `src/officecli/Core/OutputFormatter.cs`, `CommandBuilder*.cs` | JSON stdout 纯度与 stderr warning 契约 | view/get/query/batch/dump 等主要 Word JSON 链路已有独立 JSON 解析和 warning 分流；所有命令、自动更新提示和 resident busy 分支的统一纯度矩阵仍待补 | 部分自动覆盖 |

## 跨平台与发布环境

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-PLAT-001 | e2e | P1 | `.github/workflows/build.yml`, `src/officecli/officecli.csproj` | macOS/Windows 发布产物 Word 等价 smoke | `osx-arm64`、`win-x64` self-contained/single-file 分别执行 create/add/get/view/validate/close；同一 fixture 的 JSON 关键字段和 exit code 等价；禁止自动更新、自动安装和隐式 resident 影响结果 | 部分自动覆盖 |
| WORD-PLAT-002 | e2e | P1 | `src/officecli/Core/LocaleFontRegistry.cs`, `CommandBuilder.cs` | locale、UTF-8 内容和路径跨平台一致性 | `zh-CN/ar-SA` locale 的默认字体/RTL 布局、CJK/RTL 文本、非 ASCII 文件名在当前宿主发布链路可 reopen/readback/validate；Windows 发布包和 ja-JP/长路径矩阵仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |

## 本轮复审新增的命令、协议与文档矩阵

这些用例用于补齐“功能已测，但入口或组合没测”的问题。状态为“待补”表示当前没有独立自动化断言；如果已有局部覆盖，仍保留为新的细化 case，避免后续误把局部覆盖当作完整覆盖。

### 命令与跨进程协议

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-CMD-049 | e2e | P0 | `CommandBuilder.Add.cs`, `schemas/help/docx/*.json` | `add --type` 类型路由矩阵 | paragraph/run/table/picture/sdt/field/hyperlink/bookmark/header/footer 的真实 CLI 路由、query 和 validate 已固定；row/cell/section/formfield/note/comment/OLE/chart/equation/diagram/shape/textbox/watermark/permission 等类型的全量 CLI 矩阵仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-050 | e2e | P0 | `CommandBuilder.Add.cs`, `InsertPosition` | add 位置与互斥参数矩阵 | 默认 append、`--index`、`--before`、`--after`、stable-id anchor、负/越界 index、多个位置参数同时出现的 exit/code/message/文件不变性固定 | 待补 |
| WORD-CMD-051 | e2e | P0 | `CommandBuilder.Set.cs`, `CommandBuilder.Add.cs` | set/add 属性输入矩阵 | canonical key、legacy alias、大小写 key、重复 key、空值、只读 key、unsupported key、多个 `--prop` 的应用顺序和 warning/envelope 一致；支持的 key 不能因同批 unsupported key 被跳过 | 待补 |
| WORD-CMD-052 | e2e | P0 | `CommandBuilder.GetQuery.cs` | get/query path 与输出模式矩阵 | `/`、`/body`、singleton、indexed path、stable id、selector、`--depth`、`--find`、`--json`、`--save`、binary payload 和无 payload 的输出/文件创建规则固定 | 待补 |
| WORD-CMD-053 | e2e | P0 | `CommandBuilder.Add.cs`, `WordHandler.Mutations.cs` | remove/move/copy/swap 组合矩阵 | body、table row/cell、header/footer、跨 parent、同一 parent 自移动、目标为后继节点、空文档、最终 `sectPr` 保护；每个失败场景比较 body XML、package bytes、query 和 validate | 待补 |
| WORD-CMD-054 | e2e | P0 | `CommandBuilder.Check.cs`, OpenXML validator | validate 诊断矩阵 | dangling relationship、缺失 part、坏 content type、空 chart、坏 field/marker、schema error 多条并存时，exit、`success`、error 字段、part/path、排序和 stderr 形状固定 | 待补 |
| WORD-CMD-055 | integration | P1 | `CommandBuilder.Dump.cs`, `WordBatchEmitter` | dump 输入路径、格式和确定性 | `/`、`/body`、带 index 子树、无 index/不支持子树、`--out file`、`--out -`、重复 dump 的 item 顺序、warning 和裸 array/envelope 规则固定 | 待补 |
| WORD-CMD-056 | integration | P1 | `CommandBuilder.Raw.cs`, `WordHandler.Add.Media.cs` | raw/add-part relationship 矩阵 | document/settings/styles/header/footer/zip URI part 读写、chart/OLE/picture part、重复 relId、content type、query/fragment、错误 XPath/非法 XML 后原包不变 | 待补 |
| WORD-CMD-057 | e2e | P1 | `CommandBuilder.View.cs` | view 全模式与参数矩阵 | text/annotated/outline/stats/issues/forms/html/screenshot/pdf 的默认值、page/start/end/max-lines/limit/type/grid/render/out/page-count、未知 mode/非法数字、stdout 与文件输出分别固定 | 待补 |
| WORD-CMD-058 | e2e | P1 | `CommandBuilder.Watch.cs`, `WatchServer.cs` | watch HTTP/SSE 生命周期 | 启动/端口/health 或 HTML 端点、静态资源、SSE 初始事件、mutation 后版本/HTML 更新、断开重连、服务退出和临时资源清理固定 | 待补 |
| WORD-CMD-059 | e2e | P2 | `CommandBuilder.Mark.cs`, `CommandBuilder.Goto.cs` | selection/mark/goto 多目标矩阵 | paragraph/table/row/cell goto 已在真实 watch 中固定；空 selection、单/多 selection、explicit/selected mark、`--all`、unmark、stale path、无 watch 和 watch 退出后的错误 envelope 仍待补 | 部分自动覆盖 |
| WORD-CMD-060 | integration | P1 | `CommandBuilder.Save.cs`, `ResidentServer.cs`, `ResidentFlushPolicy` | resident flush policy matrix | `each`、`off`、固定秒数、`auto` 的单 mutation->内存->磁盘时序、save 保持 resident、close 释放锁已固定；idle flush 的异常保存、batch 在 off/固定/auto 下的组合和重复 close 仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-061 | integration | P0 | `ResidentClient.cs`, `ResidentServer.cs` | resident 并发与 at-most-once | 多 client 并发 add/set/remove/batch 的串行顺序、busy/retry、空响应、resident 崩溃后恢复；断线重试不能重复应用非幂等 mutation | 待补 |
| WORD-CMD-062 | integration | P1 | `CommandBuilder.Batch.cs`, `BatchExecutor` | batch 输入来源与形状矩阵 | inline、commands file、stdin、`--input -`、空 array、UTF-8 BOM、JSONL 多行当前拒绝、`--commands`/`--input` 互斥、顶层 object/array、坏 item 的主要边界和目标文件 bytes 不变已固定；stdin 重定向 warning、更多坏 item 与 stop-on-error 的输入源组合仍待补 | 部分自动覆盖 |
| WORD-CMD-063 | e2e | P1 | `CommandBuilder.Import.cs` | merge 模板和多 part token 矩阵 | body/table/header/footer/footnote/endnote/comment 的 token、数组/对象/literal key、跨 run、重复 placeholder、转义花括号、缺 key、非法 JSON、已有输出和 force 的行为固定 | 待补 |
| WORD-CMD-064 | integration | P1 | `CommandBuilder.Plugins.cs`, MCP server | plugin/MCP 错误协议矩阵 | plugin manifest 缺失/非法、unknown prop、非 dump-reader、超时/非零退出/坏 JSON/heartbeat、MCP initialize/tools/list/call/unknown tool 的 JSON-RPC 与 stderr 纯度固定 | 待补 |
| WORD-CMD-065 | e2e | P1 | `Program.cs`, `McpInstaller`, update/install guards | 自动安装、更新和环境隔离 | `OFFICECLI_NO_AUTO_INSTALL=1`、`OFFICECLI_SKIP_UPDATE=1`、无配置目录、只读目录、无网络/无插件目录时 Word 命令仍保持明确 exit/envelope，且 stdout 不混入安装日志 | 待补 |

### schema、文档对象与 round-trip

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-DOC-027 | contract | P0 | `schemas/help/docx/*.json`, `SchemaHelpRenderer` | schema/help/runtime 全量 parity | schema operations/property shape 和基础别名已由 unit 固定；所有 docx 元素的 required、enum、alias、path 与 handler 实际 add/set/get/query/remove 行为逐项对齐、发现不支持时固定 warning/error 仍待补；见 `tests/OfficeCli.Tests/Unit/WordSchemaParityTests.cs` | 部分自动覆盖 |
| WORD-DOC-028 | contract | P1 | `paragraph.json`, `run.json` | paragraph/run 高级属性矩阵 | keepNext/keepLines/pageBreakBefore/widowControl、tabs、bidi、text direction、theme/effect/shading、lang/font slot 等声明属性逐项 add/set/get/readback | 待补 |
| WORD-DOC-029 | contract | P1 | table/row/cell schemas | table complete contract | table/row/cell 的 layout/border/spacing/indent/height/header/cantSplit/merge/padding/fitText 等属性与默认值、alias、删除后的结构一致 | 待补 |
| WORD-DOC-030 | contract | P1 | picture/ole/chart/diagram/equation/shape/textbox/watermark schemas | media/visual complete contract | 每类复杂对象的必填、add-only/settable、尺寸/位置/关系、query/get、remove、HTML preview、validate 和 dump replay 一致 | 待补 |
| WORD-DOC-031 | contract | P1 | field/toc/formfield/sdt schemas | field/control complete contract | PAGE/NUMPAGES/SEQ/REF/TOC、formfield、所有 SDT type 的 instruction/result/lock/checked/list/date/placeholder/cache 行为和只读属性边界固定 | 待补 |
| WORD-DOC-032 | integration | P1 | package parts and relationship graph | unknown part preservation | 带 theme、fontTable、webSettings、customXml/customXmlProperties、embedded package、custom docProps、comments/notes/header/footer 的已有 docx 局部 mutation 后，part、rel、content type、binary bytes 和 readback 全部保留 | 待补 |
| WORD-DOC-033 | integration | P1 | `WordBatchEmitter`, `BatchExecutor` | full dump replay idempotency | full dump -> new docx -> dump 的关键 item/资源顺序、stable IDs、关系和 validate 结果可重复；第二次 replay 不产生重复 section/part/relationship | 待补 |
| WORD-DOC-034 | integration | P1 | `examples/word/*` | 示例全量 smoke 分层 | 每个 shell/Python 示例至少执行一次；每类示例断言一个专属结果（paragraph/style/table/media/field/SDT/revision/section），失败时保留 stdout/stderr 和生成文件 | 待补 |
| WORD-DOC-035 | contract | P2 | `WordHandler.HtmlPreview.*`, rendering registry | HTML/render contract | HTML 的 path/anchor/class、文本转义、SVG/data URI、table grid、field/revision/comment/note marker、page/grid 输出和 renderer capability 缺失时的降级固定 | 待补 |
| WORD-DOC-036 | contract | P1 | revision/permission/protection handlers | revision、permission、protection 交叉行为 | tracked add/set/remove/move、accept/reject、permission range、document protection、force、allowed editing range 的组合结果、marker 配对和 validate 固定 | 待补 |

### 负向、资源和平台补充矩阵

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-NEG-018 | contract | P0 | `CommandBuilder.*`, `MutationSelectorGuard` | CLI 参数和 selector 解析失败 | 缺 file/path/type、互斥 index/before/after、空 prop、坏 bool/number/length/color、bare selector、未闭合 selector、shell quote/path bracket 的 code、stderr/stdout、文件不变性固定 | 待补 |
| WORD-NEG-019 | integration | P0 | all mutation commands | mutation 全量原子性 | add/set/remove/move/copy/swap/raw-set/add-part/merge/batch 每种失败均比较原始 bytes、主体 XML、parts/rels、query、validate 和输出 envelope | 待补 |
| WORD-NEG-020 | integration | P0 | relationship cleanup helpers | relationship 冲突与删除安全 | duplicate relId、同 part 多引用、删除最后引用、删除非最后引用、跨 host part copy/remove、失败回滚时不能悬空或误删 | 待补 |
| WORD-NEG-021 | integration | P1 | document protection and permission ranges | protection 与 allowed range | protected doc 默认拒绝/force 绕过、allowed editing range 内外、batch、resident、standalone remove/raw-set 的当前差异全部固定，不能只测一个 gate | 待补 |
| WORD-NEG-022 | e2e | P0 | `ResidentClient`, `ResidentServer` | busy/crash/stale marker 矩阵 | main pipe busy、ping pipe 可用性、resident 异常退出、锁文件/marker 清理、端口/pipe 重用、下一次 open/create 的恢复和错误建议固定 | 待补 |
| WORD-NEG-023 | e2e | P0 | `OutputFormatter`, all Word commands | JSON stdout purity | 所有 `--json` Word command 的 stdout 均可单独 parse；warning/diagnostic/progress/auto-update/plugin 日志只在 stderr；裸文本模式不被 JSON envelope 污染 | 待补 |
| WORD-NEG-024 | integration | P1 | `ResidentClient.MaxMessageLength`, render/dump | 大输入和边界资源 | 超长 text、超多 batch items、超大 binary、超深 get、超多 HTML lines 的当前限制、错误和文件不变性固定；不因 pipe truncation 被误判为成功 | 待补 |
| WORD-NEG-025 | e2e | P1 | `CommandBuilder` file handling | 文件系统安全边界 | 相对/绝对/空格/CJK/RTL/`.`/`..`/符号链接/已有输出/同文件输入输出/只读目录的实际行为固定；测试不得写出临时目录 | 待补 |
| WORD-NEG-026 | integration | P1 | retry/close/save/batch semantics | 重复执行与幂等边界 | 重复 close/save/validate、同一 batch 重放、命令在响应丢失后的用户重试、open 已有 resident、create force 已占用的结果和不重复 mutation 固定 | 待补 |
| WORD-MEDIA-028 | regression | P1 | `OleHelper`, `WordHandler.Add.Media` | OLE/embedded package full round-trip | package/legacy object 两种 OLE 的 content type、progId、binary payload、display/size、dump/batch、copy/remove 后关系与 validate 固定 | 待补 |
| WORD-MEDIA-029 | regression | P1 | chart builders and resources | chart externalData/workbook round-trip | chart cache、series/category refs、embedded workbook、externalData、style/color、chart relId 和 host part 复制/删除后的完整关系固定 | 待补 |
| WORD-MEDIA-030 | regression | P1 | complex carrier handlers | ActiveX/VML/diagram/equation/rich SDT carrier | ActiveX/VML control、textbox/shape、diagram rel、equation OLE、rich SDT、embedded font、numPicBullet 等 carrier dump replay 后仍可打开、query、validate | 待补 |
| WORD-MEDIA-031 | regression | P1 | resource emitter and package graph | resource ordering and content types | full dump 的 theme/styles/numbering/settings/docProps/customXml/fontTable/webSettings/embedded parts 先后、warning、content type 和 replay 结果固定 | 待补 |
| WORD-REF-025 | e2e | P1 | `CommandBuilder.Refresh.cs`, `WordHtmlRefresh` | refresh field result matrix | TOC、PAGEREF、PAGE、NUMPAGES、SEQ、REF 分别构造前后文档，比较 result text、field cache、dirty/updateFields、Pages 和 backend/fallback envelope | 待补 |
| WORD-REF-026 | contract | P1 | `sdt.json`, `formfield.json` | controls complete matrix | text/check/dropdown/combobox/date/picture/group/richtext 的 value/result/checked/lock/placeholder/list item/default state、set/remove/dump replay 固定 | 待补 |
| WORD-REF-027 | contract | P1 | `hyperlink.json`, `WordHandler` | hyperlink target matrix | external URL、fragment anchor、既有 bookmark、theme/inherit color、run formatting、set/remove、invalid target 和关系清理固定 | 待补 |
| WORD-REF-028 | contract | P1 | comment/note handlers | comment/note range and format matrix | comment thread/reply/done/parentId/range marker、footnote/endnote reference、direction/format、跨 part dump replay、删除后 marker/part 清理固定 | 待补 |
| WORD-PLAT-003 | e2e | P0 | publish workflow and `WordPublishedBinarySmokeTests` | publish artifact execution matrix | `osx-arm64`、`win-x64` 的 single-file/self-contained 产物分别在目标 OS 启动；权限、依赖、临时目录、create/add/view/validate/close、exit code 和 JSON 关键字段一致 | 待补 |
| WORD-PLAT-004 | e2e | P1 | installer/update/runtime guards | bundled-agent isolation | 嵌入另一个 agent 时禁用 auto-install/update、默认工作目录、stdin/stdout/stderr、临时文件、resident pipe、退出码和相对路径均不污染宿主 agent | 待补 |
| WORD-PLAT-005 | e2e | P1 | refresh/render backends | native Word 与 fallback | Windows + Word native、Windows 无 Word、macOS/Linux HTML fallback 的 refresh/screenshot/pdf 能力、错误码、页面数和输出文件规则分别固定 | 待补 |
| WORD-PLAT-006 | e2e | P1 | locale/path runtime | locale/path matrix | zh-CN/ja-JP/ar-SA/en-US、CJK/RTL/emoji 文本、非 ASCII/空格/长路径、Windows separator 与 macOS separator 在发布包中 readback/validate 一致 | 待补 |

## 二次复审新增的运行时与基础设施矩阵

上一轮主要围绕 WordHandler、schema 和 CLI 主链路拆分。本轮再沿着 CodeGraph 的调用关系检查后，发现 watch runtime、打开修复、资源限制、编号渲染、字段刷新 backend 和 dump filter 仍有独立的可观察契约，却没有自己的冻结点。下面的条目专门补齐这些“主功能已测，但支撑层漂移会改变用户可见结果”的场景。

### 运行时与渲染 helper

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-UNIT-117 | unit | P0 | `src/officecli/Core/Watch/WatchNotifier.cs` | watch path/selector 转换 | paragraph、table、row、cell 的 canonical path、alias path、非法 path 分别映射到稳定 selector；不支持类型返回 null，不生成错误 selector；见 `tests/OfficeCli.Tests/Unit/WordWatchModelTests.cs` | 已有自动覆盖 |
| WORD-UNIT-118 | unit | P0 | `src/officecli/Core/Watch/WatchMark.cs`, `WatchNotifier.cs` | mark wire model 与版本 | 已固定 `WatchMark`/`MarksResponse` 的 JSON 字段、`version` 和空列表语义；`MarkRequest`/`UnmarkRequest`、重复 mark、removed count、monotonic version、无 watch 与 server reject 仍待补；见 `tests/OfficeCli.Tests/Unit/WordWatchModelTests.cs` | 部分自动覆盖 |
| WORD-UNIT-119 | unit | P1 | `src/officecli/Core/ResidentFlushPolicy.cs` | flush policy parser 与 adaptive EMA | `off`、`each`、`auto`、固定秒数、大小写/空白、越界和非法值，以及 EMA 上升/下降和 interval 上下限已固定；见 `tests/OfficeCli.Tests/Unit/WordRuntimePolicyTests.cs` | 已有自动覆盖 |
| WORD-UNIT-120 | unit | P0 | `src/officecli/Core/DocumentLimits.cs`, `DocumentHandlerFactory.cs` | OOXML zip 资源保护 | `DocumentLimits` 的 entry/size/ratio/depth/regex 常量和 `EnsureDepth` 边界、`max_depth_exceeded` 错误已固定；zip bomb、0-byte 和 repair 后文件不变性仍待补；见 `tests/OfficeCli.Tests/Unit/WordResourceLimitTests.cs` | 部分自动覆盖 |
| WORD-UNIT-121 | unit | P1 | `src/officecli/Core/WordNumFmtRenderer.cs` | Word 编号格式渲染矩阵 | 已固定 decimal/roman/letter/ordinal/cardinal/CJK、enclosed/fullwidth、Arabic/Hebrew/Thai/Devanagari、kana/Korean、干支、Chicago、hex/bullet/none 的代表值、0/负数、越界和未知格式 fallback；仍需补齐剩余 ECMA 别名和 overflow 组合；见 `tests/OfficeCli.Tests/Unit/WordNumberingAndLocaleTests.cs` | 部分自动覆盖 |
| WORD-UNIT-122 | unit | P1 | `src/officecli/Core/WordPageDefaults.cs`, `LocaleFontRegistry.cs` | 页面默认值与 locale 默认字体 | 已固定 `zh-CN`/`zh-TW`/`ar-SA` 的字体、RTL、CJK CSS fallback、字体名检测、`en-US` 对照、page dimension 边界和 `C`/`POSIX`/explicit effective-locale 优先级；更多 locale 与创建文档后的 snapshot 仍待补；见 `tests/OfficeCli.Tests/Unit/WordNumberingAndLocaleTests.cs` | 部分自动覆盖 |
| WORD-UNIT-123 | unit | P1 | `src/officecli/Core/WordTocBuilder.cs` | TOC 解析与重建 | 已固定 heading style level 过滤、`\o` level 范围、`\h` hyperlink、`\n` 无页码、PAGEREF 默认生成、旧 bookmark 复用和无 TOC 时 bookmark 生成；outline level/style definitions、跨 paragraph field、旧 entry 完整替换和更多 switch 仍待补；见 `tests/OfficeCli.Tests/Unit/WordTocBuilderTests.cs` | 部分自动覆盖 |
| WORD-UNIT-124 | unit | P1 | `src/officecli/Core/WordHtmlRefresh.cs` | refresh backend 生命周期 | backend 选择、能力缺失、超时、临时 PDF/图片清理、refresh 成功/降级/失败的字段 cache、dirty/updateFields 和错误 envelope 固定 | 待补 |
| WORD-UNIT-125 | unit | P1 | `src/officecli/Core/WordPdfBackend.cs` | PDF page filter 与图片拼接 | 单页/多页/范围/逗号列表/非法 page filter、空页、缩放、grid stitch 和 native backend 不可用时的结果与资源释放固定 | 部分自动覆盖 |
| WORD-UNIT-126 | unit | P1 | `src/officecli/Core/SchemaKeyNormalizer.cs`, `SchemaOrder.cs` | schema key alias 与 XML 顺序 | canonical case/punctuation、paragraph property schema order、detached child no-op 已固定；legacy alias/重复 key 优先级、更多容器和 Dictionary 顺序稳定性仍待补；见 `tests/OfficeCli.Tests/Unit/WordSchemaOrderTests.cs` | 部分自动覆盖 |
| WORD-UNIT-127 | unit | P1 | `src/officecli/Handlers/Word/WordBatchEmitter.Filters.cs` | dump filter 与 synthetic property 过滤 | paragraph dump 已验证过滤 `paraId/textId/relId/styleId/styleName/effective.*`；internal marker、field carrier、border/shading/padding/theme fold 和 warning 顺序仍待补；见 `tests/OfficeCli.Tests/Unit/WordDumpFilterTests.cs` | 部分自动覆盖 |
| WORD-UNIT-128 | unit | P1 | `src/officecli/Handlers/Word/WordHandler.View.cs` | view text 语义提取 | 普通 run、OLE、公式、formfield、SDT、动态 field sentinel、revision wrapper、list prefix、start/end/max-lines 截断的文本和 path 规则固定 | 部分自动覆盖 |
| WORD-UNIT-129 | unit | P2 | `src/officecli/Handlers/Rendering/BasicRenderers.cs`, `RendererRegistry` | 内置 Word renderer capability | 已固定 Word renderer 的 format/output/watch capability、错误 input 类型、重复注册幂等、可用性 fallback 和优先级选择；外部 renderer composition、真实 HTML parity 和更多 output/mode 错误输入仍待补；见 `tests/OfficeCli.Tests/Unit/WordRendererRegistryTests.cs` | 部分自动覆盖 |

### 命令入口与跨进程补充

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-CMD-066 | e2e | P1 | `CommandBuilder.Mark.cs`, `WatchNotifier.cs` | mark/list/get-marks/unmark CLI 协议 | 已覆盖 explicit path、regex find、bare hex color 归一化、非法 color reject、`get-marks` version/JSON、按 path 和 `--all` unmark；重复 mark、无 watch、完整 server reject envelope、selector 输入和 stdout/stderr 纯度仍待补；见 `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 部分自动覆盖 |
| WORD-CMD-067 | e2e | P1 | `CommandBuilder.Watch.cs`, `WatchNotifier.cs`, `WatchServer.cs` | selection 与 SSE 断开重连 | 空/单/多 selection、mutation 后 version 递增、HTML 更新事件、断开重连、watch close 后查询结果和 pipe 清理固定 | 部分自动覆盖 |
| WORD-CMD-068 | integration | P0 | `DocumentHandlerFactory.cs`, `CommandBuilder.cs` | open/repair/corrupt 输入矩阵 | 0-byte、非 zip、坏 XML encoding、dangling internal relationship、坏 content type、重试打开分别返回稳定 code；repair 后原文件可 reopen 且不误删有效 relationship | 部分自动覆盖 |
| WORD-CMD-069 | e2e | P1 | `CommandBuilder.cs` | create 已存在目标与 force 矩阵 | 已有文件、已有 resident、扩展名/type 不一致、`--force`、同一输入输出路径的 exit/code/文件 bytes/锁清理固定 | 部分自动覆盖 |
| WORD-CMD-070 | e2e | P1 | `CommandBuilder.View.cs`, rendering registry | view renderer 选择与输出冲突 | html/svg/screenshot/pdf 的 backend capability、`--render`、`--out`/stdout、已有输出、page filter/grid 和未知 renderer 的 JSON/裸文本错误边界固定 | 部分自动覆盖 |
| WORD-CMD-071 | e2e | P1 | `CommandBuilder.Refresh.cs`, `WordHtmlRefresh.cs` | refresh CLI page/filter/backend 矩阵 | `--fields`、`--pages`、`--timeout`、native/fallback/no-backend、输出文件和 JSON envelope 的组合固定；失败不覆盖原文件 | 部分自动覆盖 |

### 文档对象与跨层 round-trip 补充

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-DOC-037 | contract | P1 | `WordHandler.View.cs`, field/form/SDT/OLE/equation handlers | Word 文本视图语义矩阵 | 同一 fixture 在 text/annotated/outline/stats/issues/forms 中对 paragraph/table/field/formfield/SDT/OLE/equation/revision 的 path、label、text、统计数量保持可解释一致 | 部分自动覆盖 |
| WORD-DOC-038 | contract | P1 | `WordNumFmtRenderer.cs`, numbering schemas | 编号显示与 round-trip | list style/abstractNum/num 的 `numFmt` 全量代表值、start、restart、级别继承和 view text marker 与 dump replay 结果一致 | 部分自动覆盖 |
| WORD-DOC-039 | contract | P1 | `WordHandler.Add.Structure.cs`, header/footer helpers | section/header/footer/watermark 关系生命周期 | 多 section 的 link-to-previous、titlePg/evenAndOdd、header/footer/watermark add/set/remove、最后引用删除和 reopen 后关系图一致 | 部分自动覆盖 |
| WORD-DOC-040 | contract | P1 | `WordHandler.Add.Structure.cs`, style/numbering helpers | style/numbering 继承与 ID 冲突 | style basedOn/next、默认 style、num/abstractNum/level 继承、重复 styleId/numId、删除被引用项的当前错误或保留策略固定 | 部分自动覆盖 |
| WORD-DOC-041 | contract | P1 | `WordHandler.Add.Misc.cs`, `WordHandler.HtmlPreview.Shapes.cs` | textbox/shape/watermark 参数清洗 | geometry、wrap、anchor、rotation、fill/line/shadow、text direction、opacity、alt text 的 alias、非法值 fallback、HTML/readback/validate 一致 | 部分自动覆盖 |
| WORD-DOC-042 | integration | P1 | `WordBatchEmitter.Filters.cs`, package parts | auxiliary part 与 dump filter parity | full/subtree dump 对 theme/styles/numbering/settings/docProps/customXml/fontTable/webSettings、field carrier、border/shading/theme 的 emit/filter/replay 顺序与 warning 完整固定 | 部分自动覆盖 |

### 负向、资源和重试补充

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-NEG-027 | integration | P0 | `DocumentHandlerFactory.cs`, `DocumentLimits` | 打开修复与资源限制失败原子性 | encoding repair、dangling rel repair、坏 zip、0-byte、decompression bomb 的 exit/code、原始 bytes、临时文件和后续 reopen 结果固定 | 部分自动覆盖 |
| WORD-NEG-028 | e2e | P1 | `WatchNotifier.cs`, `CommandBuilder.Mark.cs` | mark no-watch 与 rejected 请求区分 | no watch 返回“start watch”类错误；watch 存在但 path/color/regex 被拒绝返回真实 reject；空 id、空 error、超时不会伪装成成功 | 部分自动覆盖 |
| WORD-NEG-029 | e2e | P1 | `CommandBuilder.*`, file handling | 输出路径、临时文件和同路径边界 | `--out` 已存在、输入输出相同、父目录不存在、只读目录、符号链接、`.`/`..`、进程中断后的临时文件清理和源文件不变固定 | 待补 |
| WORD-NEG-030 | integration | P0 | relationship helpers, raw/add-part | relationship 冲突与回滚矩阵 | duplicate relId、相同 part 多引用、删除最后/非最后引用、坏 content type、raw-set/add-part/mutation 失败回滚时不能误删或留下 dangling rel | 部分自动覆盖 |
| WORD-NEG-031 | integration | P0 | `ResidentClient.cs`, `ResidentServer.cs`, `ResidentFlushPolicy.cs` | timeout、半响应与 at-most-once | server 中途断开、响应为空/延迟、client retry、save/close 超时、非幂等 mutation 的发送次数和最终文档顺序固定 | 部分自动覆盖 |

## 首批冻结基线用例

这批用例优先用于锁定当前行为。落地后再扩展完整清单。

| ID | 来源 | 层级 | 建议测试文件 | 冻结行为 |
|---|---|---|---|---|
| WORD-FREEZE-001 | WORD-UNIT-008 | unit | `tests/OfficeCli.Tests/Unit/WordParagraphTests.cs` | paragraph add/set/get 的 `align/spaceBefore/spaceAfter/lineSpacing` canonical readback |
| WORD-FREEZE-002 | WORD-UNIT-009 | unit | `tests/OfficeCli.Tests/Unit/WordRunTests.cs` | run 文本、bold、italic、font、size、color、underline readback |
| WORD-FREEZE-003 | WORD-UNIT-010 | unit | `tests/OfficeCli.Tests/Unit/WordSelectorTests.cs` | selector 支持 child combinator、`:contains()`、`:empty`、大小写宽容 |
| WORD-FREEZE-004 | WORD-UNIT-001 | unit | `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | inline picture query 返回 `wrap=inline` |
| WORD-FREEZE-005 | WORD-UNIT-002 | unit | `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | anchored picture query 返回 wrap、position、relative frame |
| WORD-FREEZE-006 | WORD-UNIT-003 | unit | `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | OLE query 与 picture query 不互相污染 |
| WORD-FREEZE-007 | WORD-REF-013 | contract | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | SDT schema 只声明已实现 add-time type |
| WORD-FREEZE-008 | WORD-REF-014 | contract | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | SDT `type` 是 add/get canonical key，`sdtType` 不作为 readback key |
| WORD-FREEZE-009 | WORD-REF-015 | contract | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | SDT lock 派生 editable，并在重开后保持 |
| WORD-FREEZE-010 | WORD-REF-016 | contract | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | checkbox SDT readback `checked` |
| WORD-FREEZE-011 | WORD-CMD-001 | integration | `tests/OfficeCli.Tests/Integration/WordCliSmokeTests.cs` | CLI `create -> add paragraph -> get -> close` 最小链路 |
| WORD-FREEZE-012 | WORD-CMD-002 | integration | `tests/OfficeCli.Tests/Integration/WordCliSmokeTests.cs` | CLI `set` 后 `get` 和 `validate` 看到最新行为 |
| WORD-FREEZE-013 | WORD-CMD-003 | integration | `tests/OfficeCli.Tests/Integration/WordQueryTests.cs` | `query` 返回稳定路径、数量、空结果 |
| WORD-FREEZE-014 | WORD-CMD-006 | integration | `tests/OfficeCli.Tests/Integration/WordBatchTests.cs` | batch 多步结果与单命令链路一致 |
| WORD-FREEZE-015 | WORD-CMD-007 | integration | `tests/OfficeCli.Tests/Integration/WordDumpBatchRoundTripTests.cs` | 小型 `dump -> batch` 后关键节点可查询且 `validate` 通过 |
| WORD-FREEZE-016 | WORD-CMD-016 | integration | `tests/OfficeCli.Tests/Integration/WordRawTests.cs` | `raw` 可读取 document/settings/styles part，错误 part 有结构化错误 |
| WORD-FREEZE-017 | WORD-CMD-017 | integration | `tests/OfficeCli.Tests/Integration/WordRawTests.cs` | `raw-set` 成功写回；错误 XPath 不污染文件 |
| WORD-FREEZE-018 | WORD-TBL-003 | contract | `tests/OfficeCli.Tests/Functional/WordTableContractTests.cs` | cell text、width、shd、align、valign、padding readback |
| WORD-FREEZE-019 | WORD-TBL-004 | contract | `tests/OfficeCli.Tests/Functional/WordTableContractTests.cs` | `gridspan/hmerge/vmerge` 后索引和吸收 cell 行为 |
| WORD-FREEZE-020 | WORD-DOC-016 | contract | `tests/OfficeCli.Tests/Functional/WordSectionContractTests.cs` | section length/orientation/margin/columns readback |
| WORD-FREEZE-021 | WORD-REV-001 | contract | `tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | run-level revision query 返回 type、author、date、path |
| WORD-FREEZE-022 | WORD-REV-002 | contract | `tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | paragraph add/remove/set with revision.author 的当前 marker 行为 |
| WORD-FREEZE-023 | WORD-MEDIA-020 | regression | `tests/OfficeCli.Tests/Integration/WordDumpBatchRoundTripTests.cs` | 图片 round-trip 后 part/rel 不丢失 |
| WORD-FREEZE-024 | WORD-MEDIA-021 | regression | `tests/OfficeCli.Tests/Integration/WordDumpBatchRoundTripTests.cs` | OLE / embedded package round-trip 后 part/rel 不丢失 |
| WORD-FREEZE-025 | WORD-MEDIA-023 | regression | `tests/OfficeCli.Tests/Integration/WordDumpBatchRoundTripTests.cs` | header/footer 内媒体 rel 挂在正确 host part |
| WORD-FREEZE-026 | WORD-NEG-001 | contract | `tests/OfficeCli.Tests/Functional/WordErrorContractTests.cs` | 非法 path 返回结构化错误且文件不变 |
| WORD-FREEZE-027 | WORD-NEG-002 | contract | `tests/OfficeCli.Tests/Functional/WordErrorContractTests.cs` | 非法 enum/bool/length/color 返回结构化错误和可用提示 |
| WORD-FREEZE-028 | WORD-NEG-003 | contract | `tests/OfficeCli.Tests/Functional/WordErrorContractTests.cs` | 缺 required prop 不创建半成品节点或 dangling rel |
| WORD-FREEZE-029 | WORD-NEG-007 | integration | `tests/OfficeCli.Tests/Integration/WordBatchTests.cs` | batch 默认继续，`--stop-on-error` 停止 |
| WORD-FREEZE-030 | WORD-CMD-010 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `view text/annotated/outline/stats/issues --json` 输出结构稳定 |
| WORD-FREEZE-031 | WORD-UNIT-086 | unit | `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | chart add/get/query 基础 readback 和 `validate` 空结果 |
| WORD-FREEZE-032 | WORD-UNIT-087 | unit | `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | chart series/axis set、range refs/color 和非法数值当前错误边界 |
| WORD-FREEZE-033 | WORD-UNIT-088 | unit | `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | native diagram group/textbox、resize/remove、缺 source 不污染 |
| WORD-FREEZE-034 | WORD-UNIT-089 | unit | `tests/OfficeCli.Tests/Unit/WordChartDiagramPreviewTests.cs` | HTML preview 段落、表格、图片 data URI 和 alt 文本 |
| WORD-FREEZE-035 | WORD-CMD-021/022/024/025/027 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 真实 CLI `create/add/get/set/query/remove/validate` 最小闭环，含 `create --force` 和 `add` 位置/复制规则 |
| WORD-FREEZE-036 | WORD-CMD-023/WORD-MEDIA-024 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `get --save` 提取 picture/OLE payload 和无 payload 失败 |
| WORD-FREEZE-037 | WORD-CMD-012/029 | integration/e2e | `tests/OfficeCli.Tests/Integration/WordResidentLifecycleTests.cs`, `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | resident save/close 的磁盘可见性，以及无 resident 时 `save` no-op 成功 |
| WORD-FREEZE-038 | WORD-UNIT-090/091/092 | unit | `tests/OfficeCli.Tests/Unit/WordRunTests.cs` | run 高级格式、text effects、theme color 当前 readback |
| WORD-FREEZE-039 | WORD-DOC-022 | contract | `tests/OfficeCli.Tests/Functional/WordSchemaHelpParityTests.cs` | schema/help/runtime 不漂移，alias 写入 canonical readback |
| WORD-FREEZE-040 | WORD-NEG-010 | integration | `tests/OfficeCli.Tests/Integration/WordProtectionGateTests.cs` | protected docx 默认拒绝 mutation，`--force` 绕过 |
| WORD-FREEZE-041 | WORD-CMD-039 | e2e | `tests/OfficeCli.Tests/E2E/WordPublishedBinarySmokeTests.cs` | 发布后的 macOS/Windows 二进制可执行最小 Word 链路且无安装/更新副作用 |
| WORD-FREEZE-042 | WORD-UNIT-097 | unit | `tests/OfficeCli.Tests/Unit/WordTextInputBoundaryTests.cs` | 换行/tab/page field token/非法 XML 字符的当前文本输入行为 |
| WORD-FREEZE-043 | WORD-MEDIA-027 | regression | `tests/OfficeCli.Tests/Integration/WordComplexDumpCarrierRoundTripTests.cs` | ActiveX/VML/复杂 carrier dump replay 后关系不悬空 |
| WORD-FREEZE-044 | WORD-UNIT-098 | unit | `tests/OfficeCli.Tests/Unit/WordBreakTabPermissionTests.cs` | tab 负 position、非法 enum、header/footer ptab 格式 readback |
| WORD-FREEZE-045 | WORD-UNIT-096 | unit | `tests/OfficeCli.Tests/Unit/WordMediaQueryTests.cs` | picture/OLE binary payload extraction 和无 payload 不落空文件 |
| WORD-FREEZE-046 | WORD-UNIT-094 | unit | `tests/OfficeCli.Tests/Unit/WordStableIdTests.cs` | editable open 稳定 ID 补齐/去重，read-only open 不写回 |
| WORD-FREEZE-047 | WORD-UNIT-093 | unit | `tests/OfficeCli.Tests/Unit/WordRunTests.cs` | range 半开区间 run split/format 和互斥错误 |
| WORD-FREEZE-048 | WORD-UNIT-095 | unit | `tests/OfficeCli.Tests/Unit/WordStrictAttributeSanitizerTests.cs` | strict 属性清理和失败打开释放文件句柄 |
| WORD-FREEZE-049 | WORD-NEG-012 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI bare `key=value` 缺 `--prop` 的 warning、exit code 和不应用属性行为 |
| WORD-FREEZE-050 | WORD-NEG-010 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | docx protection gate 当前命令面：add/set/batch 拦截与 force 绕过，standalone remove/raw-set 当前绕过 |
| WORD-FREEZE-051 | WORD-CMD-031 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `dump --out file` 写裸 JSON array、`--out -` 回 stdout、invalid format error code |
| WORD-FREEZE-052 | WORD-CMD-032 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `raw/raw-set/add-part` CLI 基础协议和当前 chart part validation 状态 |
| WORD-FREEZE-053 | WORD-NEG-011/WORD-CMD-026 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI 缺文件/不支持扩展/路径越界/无 watch selected/bare selector 错误 envelope |
| WORD-FREEZE-054 | WORD-CMD-015/WORD-CMD-037 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `merge` docx body/table/header/footer 占位符替换、未命中 key、数据 flatten、`file_exists` 和 `invalid_json` 当前行为 |
| WORD-FREEZE-055 | WORD-CMD-030 | integration | `tests/OfficeCli.Tests/Integration/WordResidentLifecycleTests.cs` | resident batch 内 `open/close` 跳过、会话保持存活、batch 内 query 可见新增内容、显式 save 后磁盘可见 |
| WORD-FREEZE-056 | WORD-CMD-041 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `swap` CLI body 段落和表格 cell 交换、跨父级 `invalid_value`、缺失路径 `not_found` 当前行为 |
| WORD-FREEZE-057 | WORD-REF-020 | contract | `tests/OfficeCli.Tests/Functional/WordHyperlinkContractTests.cs` | hyperlink wrapper metadata、run formatting、theme color 和 inherit sentinel readback 当前行为 |
| WORD-FREEZE-058 | WORD-DOC-021 | contract | `tests/OfficeCli.Tests/Functional/WordFindReplaceContractTests.cs` | find/regex replace、range 跨段落半开区间格式化、互斥和格式限定错误当前行为 |
| WORD-FREEZE-059 | WORD-REF-021 | contract | `tests/OfficeCli.Tests/Functional/WordNotesAndCommentsContractTests.cs` | comment 扩展元数据、done/parentId 回复链、rangeOpen/rangeEnd 删除清理当前行为 |
| WORD-FREEZE-060 | WORD-REF-022 | contract | `tests/OfficeCli.Tests/Functional/WordNotesAndCommentsContractTests.cs` | footnote/endnote 格式 add/get readback 和 remove 清理当前行为 |
| WORD-FREEZE-061 | WORD-CMD-013 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `OFFICECLI_RESIDENT_FLUSH=each` 下 CLI mutation 返回即磁盘可见 |
| WORD-FREEZE-062 | WORD-CMD-025 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `set` CLI 缺属性、裸属性、unsupported warning/envelope 当前行为 |
| WORD-FREEZE-063 | WORD-CMD-027 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `remove/move/copy` CLI 路径、位置参数、输出和段落顺序当前行为 |
| WORD-FREEZE-064 | WORD-CMD-024 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `query` CLI child selector、boolean filter 和 hydrated JSON children 当前行为 |
| WORD-FREEZE-065 | WORD-CMD-025 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `set` CLI scoped selector 多命中后所有匹配段落的属性 readback |
| WORD-FREEZE-066 | WORD-CMD-026 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | resident 存活时 CLI scoped selector set 转发、内存 readback 与 close 后落盘 |
| WORD-FREEZE-067 | WORD-CMD-028 | integration | `tests/OfficeCli.Tests/Integration/WordResidentLifecycleTests.cs` | resident 与 non-resident batch JSON `results/summary/item error` envelope parity |
| WORD-FREEZE-068 | WORD-CMD-028 | integration | `tests/OfficeCli.Tests/Integration/WordResidentLifecycleTests.cs` | resident 与 non-resident batch `--stop-on-error` 的 executed/skipped envelope parity |
| WORD-FREEZE-069 | WORD-CMD-015/037 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | merge 对 footnote/endnote/comment 的替换，以及跨 run 占位符保留未解析的当前边界 |
| WORD-FREEZE-070 | WORD-CMD-032 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | raw zip URI 读取（query/fragment 忽略）与不存在 part 的 `invalid_value` envelope |
| WORD-FREEZE-071 | WORD-REF-021 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI 当前 `query comment[direction=rtl]` 的 `unknown_key`/空结果边界，以及无过滤 query 的格式面 |
| WORD-FREEZE-072 | WORD-UNIT-093 | unit | `tests/OfficeCli.Tests/Unit/WordRunTests.cs` | 多段逗号 range 按位置排序并只格式化各自半开区间 |
| WORD-FREEZE-073 | WORD-NEG-013 | integration | `tests/OfficeCli.Tests/Integration/WordComplexRelationshipCleanupTests.cs` | 删除 picture/OLE/hyperlink/comment/note/header/footer 后 query 清空且 validate 无错误 |
| WORD-FREEZE-074 | WORD-MEDIA-001/006 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI `query picture` 的 inline wrap 与 `query ole` 的类型/progId 区分 |
| WORD-FREEZE-075 | WORD-MEDIA-002 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI `query picture` 的 floating wrap、位置和 relative frame |
| WORD-FREEZE-076 | WORD-CMD-027 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI `/body/ole[1]` indexed shorthand 删除后 query 为空且 validate 无错误 |
| WORD-FREEZE-077 | WORD-CMD-012/029 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI `open` 自举 resident、add/query 通过 pipe、`close` 后外部重开可读 |
| WORD-FREEZE-078 | WORD-REV-001/013 | contract | `tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | revision date 归一化、缺 date 当前行为和重复 id 的 type-specific path 消歧 |
| WORD-FREEZE-079 | WORD-NEG-012 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI set 大小写属性 key、alignment legacy alias 和 unsupported warning 的当前 envelope |
| WORD-FREEZE-080 | WORD-CMD-018 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI 连续 `add-part chart` 生成不同 relationship、`/chart[1]` 与 `/chart[2]` raw 可读 |
| WORD-FREEZE-081 | WORD-CMD-041 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | resident 存活时 CLI swap 转发、内存顺序和 close 后磁盘顺序 |
| WORD-FREEZE-082 | WORD-REV-006/013 | contract | `tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | revision nativePath 动作、author/type 批量筛选、stale id 和 creation/action 混用保护当前行为 |
| WORD-FREEZE-083 | WORD-DOC-022 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `help docx revision` 的 operations/property contract、`help docx set revision` 路由和 `help docx all` 汇总当前行为 |
| WORD-FREEZE-084 | WORD-DOC-021/REV-015 | contract | `tests/OfficeCli.Tests/Functional/WordFindReplaceContractTests.cs` | `find + revision.author` 的 replacement、空 replacement 和 format-only 三种 tracked 分支当前行为 |
| WORD-FREEZE-085 | WORD-REV-012 | contract | `tests/OfficeCli.Tests/Functional/WordRevisionContractTests.cs` | `remove + revision.*` 的 run 删除 marker、paragraph/row/cell structural marker，以及 run accept/reject 当前行为 |
| WORD-FREEZE-086 | WORD-CMD-032 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI `raw-set` 非法 XML 返回 `internal_error`，且原文档内容保持不变 |
| WORD-FREEZE-087 | WORD-CMD-031 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `dump --json` 的 orphan footnote warning 同时出现在 envelope/stderr，且不污染 batch JSON data |
| WORD-FREEZE-088 | WORD-CMD-027 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | resident 中 CLI remove 转发、内存 query 和 close 后磁盘删除结果 |
| WORD-FREEZE-089 | WORD-CMD-029 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | resident 固定 1 秒 flush 下，外部 OpenXML reader 在不 save/close 时最终可见 CLI mutation |
| WORD-FREEZE-090 | WORD-CMD-036 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI refresh 的 unsupported type、refresh_failed 降级和成功 backend envelope 当前行为 |
| WORD-FREEZE-091 | WORD-REF-023 | contract | `tests/OfficeCli.Tests/Functional/WordFormFieldContractTests.cs` | dropdown 非法 text 改缓存显示、不改 result；直接 set result 返回 unsupported，文档 validate 无错误 |
| WORD-FREEZE-092 | WORD-REF-022 | contract | `tests/OfficeCli.Tests/Functional/WordNotesAndCommentsContractTests.cs` | footnote/endnote direction 的 Format readback、align 的 raw XML 写入与当前缺失 Format key 行为 |
| WORD-FREEZE-093 | WORD-MEDIA-023 | integration | `tests/OfficeCli.Tests/Integration/WordDumpBatchRoundTripTests.cs` | dump/batch round-trip 同时保留 header/footer 图片 name/alt 与有效关系 |
| WORD-FREEZE-094 | WORD-DOC-022 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `help docx all` 列出全部嵌入 docx schema 元素，且每项 operations/path 摘要完整 |
| WORD-FREEZE-095 | WORD-CMD-040 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | MCP stdio JSON-RPC 初始化、工具发现、Word create/add/query 调用和未知工具错误 envelope 当前行为 |
| WORD-FREEZE-096 | WORD-NEG-013 | integration | `tests/OfficeCli.Tests/Integration/WordComplexRelationshipCleanupTests.cs` | 删除 chart 后 part、正文引用清理，query 为空且 validate 无错误 |
| WORD-FREEZE-097 | WORD-CMD-020 | integration | `tests/OfficeCli.Tests/Integration/WordPluginContractTests.cs` | dump-reader JSONL replay 到 docx，以及顶层 JSON array 的 `corrupt_batch` 协议 |
| WORD-FREEZE-098 | WORD-CMD-038 | integration | `tests/OfficeCli.Tests/Integration/WordPluginContractTests.cs` | plugins lint 合法 docx prop 成功、JSON array 输出返回 `corrupt_batch` |
| WORD-FREEZE-099 | WORD-CMD-030 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | resident `each` flush 下 batch 返回后外部 OpenXML reader 立即可见 mutation |
| WORD-FREEZE-100 | WORD-CMD-009/032 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | CLI validate 对空 chart part 返回结构化 schema error 的当前 envelope |
| WORD-FREEZE-101 | WORD-DOC-023 | integration | `tests/OfficeCli.Tests/Integration/WordExampleScriptSmokeTests.cs` | 四个代表 Word shell 示例在临时目录执行、生成 docx 并可重新打开查询 paragraph |
| WORD-FREEZE-102 | WORD-CMD-042 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | `view forms` 的保护状态/SDT 字段 JSON，以及未知 view mode 的 `invalid_value` envelope |
| WORD-FREEZE-103 | WORD-MEDIA-025 | unit | `tests/OfficeCli.Tests/Unit/WordDumpEmitterTests.cs` | full dump 的 numbering/styles/theme/settings/core/app docProps 在 body item 前发出且无 warning |
| WORD-FREEZE-104 | WORD-REV-014 | functional | `tests/OfficeCli.Tests/Functional/WordAliasContractTests.cs` | 六类高频 legacy alias 写入后读取 canonical key，validate 无错误；picture `w/h` 不在支持断言内 |
| WORD-FREEZE-105 | WORD-CMD-039 | e2e | `tests/OfficeCli.Tests/E2E/WordPublishedBinarySmokeTests.cs` | macOS/Windows single-file 发布产物存在且当前宿主平台完成 Word create/add/view/validate smoke |
| WORD-FREEZE-106 | WORD-MEDIA-027 | integration | `tests/OfficeCli.Tests/Integration/WordComplexDumpCarrierRoundTripTests.cs` | chart inlinedparts carrier 的 host relId、chart style child、external relationship 和 XML dump→batch round-trip；fixture schema 完整性另由 chart validation 用例覆盖 |
| WORD-FREEZE-107 | WORD-MEDIA-026 | functional | `tests/OfficeCli.Tests/Functional/WordHtmlPreviewContractTests.cs` | HTML preview 的复杂对象 SVG/文本、水印 class 和 paragraph data-path 当前行为 |
| WORD-FREEZE-108 | WORD-CMD-033/034/035 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 真实 watch 生命周期中的空 selection、explicit mark/list/unmark、paragraph/table/row/cell goto 成功和缺失 anchor 错误 |
| WORD-FREEZE-109 | WORD-CMD-043/WORD-PLAT-002 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | create 的 type/locale/minimal 标准与最小 docx parts、RTL section、CJK/RTL 文本和非 ASCII 路径 |
| WORD-FREEZE-110 | WORD-CMD-044 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | get depth=0/1、超大 depth 限制、缺失 path 当前错误 envelope |
| WORD-FREEZE-111 | WORD-CMD-045 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | view text/annotated 窗口、max-lines、issues type/limit 当前输出 |
| WORD-FREEZE-112 | WORD-CMD-046/047 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | HTML stdout/out 等价、screenshot 单页/grid/native 错误边界、pdf 产物或 exporter 错误 |
| WORD-FREEZE-113 | WORD-CMD-048 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | batch inline/file/stdin/--input - 等价、坏输入 code 和目标文件不变 |
| WORD-FREEZE-114 | WORD-DOC-024 | integration | `tests/OfficeCli.Tests/Integration/WordDocumentFidelityTests.cs` | customXml、fontTable、webSettings 在局部 mutation 与 dump/batch 后保留 |
| WORD-FREEZE-115 | WORD-DOC-025 | functional | `tests/OfficeCli.Tests/Functional/WordSectionRelationshipTests.cs` | 多 section 的 header/footer reference 指向对应 part，内容归属和 validate 稳定 |
| WORD-FREEZE-116 | WORD-DOC-026/WORD-REF-024 | functional/e2e | `tests/OfficeCli.Tests/Functional/WordFieldViewContractTests.cs`, `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | field_not_evaluated/field_cache_stale、view sentinel、refresh backend/Pages 状态 |
| WORD-FREEZE-117 | WORD-NEG-014/015/016 | integration/e2e | `tests/OfficeCli.Tests/Integration/WordMutationAtomicityTests.cs`, `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | failed mutation bytes不变、resident file_locked/异常退出恢复、Unicode 路径当前链路 |
| WORD-FREEZE-118 | WORD-UNIT-099/100/101 | unit | `tests/OfficeCli.Tests/Unit/WordSchemaParityTests.cs`, `tests/OfficeCli.Tests/Unit/WordTypedAttributeFallbackTests.cs`, `tests/OfficeCli.Tests/Unit/WordPathAliasTests.cs`, `tests/OfficeCli.Tests/Unit/WordNavigationTests.cs`, `tests/OfficeCli.Tests/Unit/WordSelectorTests.cs` | schema operation/property、typed prop parser、path/selector grammar 的当前边界 |
| WORD-FREEZE-119 | WORD-UNIT-102/103/104 | unit/integration | `tests/OfficeCli.Tests/Unit/WordMutationPositionTests.cs`, `tests/OfficeCli.Tests/Integration/WordRelationshipGraphTests.cs` | 插入位置、relId 冲突、host part、content type 和删除安全 |
| WORD-FREEZE-120 | WORD-UNIT-105/106/107 | unit/functional | `tests/OfficeCli.Tests/Unit/WordFieldParserTests.cs`, `tests/OfficeCli.Tests/Unit/WordHeaderFooterTests.cs`, `tests/OfficeCli.Tests/Unit/WordDocumentSettingsTests.cs`, `tests/OfficeCli.Tests/Functional/WordFieldContractTests.cs`, `tests/OfficeCli.Tests/Functional/WordSectionRelationshipTests.cs` | dynamic/interactive field 分类、SEQ cache 重算、header/footer reference flags、section break/page/line numbering 已部分固定；linkage 继承、section 增删与 body sectPr 归属仍待补 |
| WORD-FREEZE-121 | WORD-UNIT-112/113/114 | unit/functional | `tests/OfficeCli.Tests/Unit/WordDumpEmitterTests.cs`, `tests/OfficeCli.Tests/Unit/WordHtmlSafetyTests.cs`, `tests/OfficeCli.Tests/Functional/WordHtmlPreviewContractTests.cs` | dump 顺序与连续 JSON 稳定、共享 HTML 转义与 media data URI 降级已部分固定；warning 全量矩阵、Word 专用 path/attribute 和 validator 诊断结构仍待补 |
| WORD-FREEZE-122 | WORD-CMD-049/050/051 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | 代表 Word add 类型路由、插入位置互斥、prop alias/unsupported warning 当前协议 |
| WORD-FREEZE-123 | WORD-CMD-052/053/054 | e2e | `tests/OfficeCli.Tests/E2E/WordMutationCommandMatrixTests.cs` | get/query 输出、remove/move/copy/swap 组合、validate 诊断矩阵 |
| WORD-FREEZE-124 | WORD-CMD-055/056/062 | integration | `tests/OfficeCli.Tests/Integration/WordBatchProtocolTests.cs` | dump/raw/add-part、输入来源/空数组/BOM/JSONL 和 replay envelope |
| WORD-FREEZE-125 | WORD-CMD-057/058/059 | e2e | `tests/OfficeCli.Tests/E2E/WordViewWatchProtocolTests.cs` | view 全模式、watch HTTP/SSE、selection/mark/goto 多目标 |
| WORD-FREEZE-126 | WORD-CMD-060/061 | integration | `tests/OfficeCli.Tests/Integration/WordResidentProtocolTests.cs` | flush policy、并发顺序、busy/crash 后恢复与 at-most-once |
| WORD-FREEZE-127 | WORD-CMD-063/064/065 | integration/e2e | `tests/OfficeCli.Tests/Integration/WordPluginProtocolTests.cs`, `tests/OfficeCli.Tests/E2E/WordEnvironmentIsolationTests.cs` | merge token、plugin/MCP 错误、自动安装/更新和 stdout 隔离 |
| WORD-FREEZE-128 | WORD-DOC-027/028/029/030 | contract | `tests/OfficeCli.Tests/Functional/WordSchemaRuntimeParityTests.cs` | schema/help/runtime 全量 parity 与 paragraph/run/table/media 属性矩阵 |
| WORD-FREEZE-129 | WORD-DOC-031/032/033 | integration | `tests/OfficeCli.Tests/Integration/WordFieldAndPackageRoundTripTests.cs` | field/control、未知 package part、full dump replay 幂等性 |
| WORD-FREEZE-130 | WORD-DOC-034/035/036 | integration/contract | `tests/OfficeCli.Tests/Integration/WordExampleMatrixTests.cs`, `tests/OfficeCli.Tests/Functional/WordRenderingContractTests.cs` | 示例分层 smoke、HTML/render、revision/permission/protection 交叉行为 |
| WORD-FREEZE-131 | WORD-NEG-018/019/020 | contract/integration | `tests/OfficeCli.Tests/Integration/WordFailureAtomicityMatrixTests.cs` | CLI 参数失败、全 mutation 原子性、relationship 冲突/清理安全 |
| WORD-FREEZE-132 | WORD-NEG-021/022/023 | integration/e2e | `tests/OfficeCli.Tests/Integration/WordProtectionResidentTests.cs`, `tests/OfficeCli.Tests/E2E/WordJsonPurityTests.cs` | protection range、resident busy/crash/stale marker、JSON stdout/stderr 纯度 |
| WORD-FREEZE-133 | WORD-NEG-024/025/026 | integration/e2e | `tests/OfficeCli.Tests/Integration/WordResourceLimitTests.cs`, `tests/OfficeCli.Tests/E2E/WordFilesystemBoundaryTests.cs` | 大输入、文件系统边界、重复执行/重试不重复 mutation |
| WORD-FREEZE-134 | WORD-MEDIA-028/029/030/031 | regression | `tests/OfficeCli.Tests/Integration/WordMediaCarrierRoundTripTests.cs` | OLE/chart workbook/ActiveX/VML/complex carrier/resource ordering round-trip |
| WORD-FREEZE-135 | WORD-REF-025/026/027/028 | functional/e2e | `tests/OfficeCli.Tests/Functional/WordReferenceMatrixTests.cs`, `tests/OfficeCli.Tests/E2E/WordRefreshMatrixTests.cs` | refresh 字段结果、controls、hyperlink、comment/note range 和关系清理 |
| WORD-FREEZE-136 | WORD-PLAT-003/004/005/006 | e2e | `tests/OfficeCli.Tests/E2E/WordPlatformMatrixTests.cs` | macOS/Windows 发布包、agent 隔离、native/fallback、locale/path 等价性 |
| WORD-FREEZE-137 | WORD-CMD-027/WORD-NEG-014 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | resident move/copy 的内存顺序、显式 `save` 成功和 close 后外部重排顺序 |
| WORD-FREEZE-138 | WORD-CMD-048/WORD-CMD-062 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | batch JSONL 当前 `invalid_json`、commands/input 冲突当前 `internal_error`、失败前目标 bytes 不变 |
| WORD-FREEZE-139 | WORD-CMD-029/WORD-CMD-060 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | resident adaptive auto flush 在 idle 窗口内将 mutation 写入磁盘 |
| WORD-FREEZE-140 | WORD-UNIT-117/WORD-UNIT-118 | unit | `tests/OfficeCli.Tests/Unit/WordWatchModelTests.cs` | watch path selector、mark wire model、no-watch/reject 区分和 mark version |
| WORD-FREEZE-141 | WORD-UNIT-119/WORD-UNIT-120 | unit | `tests/OfficeCli.Tests/Unit/WordRuntimePolicyTests.cs` | resident flush policy 边界、adaptive interval、OOXML zip resource guard |
| WORD-FREEZE-142 | WORD-UNIT-121/WORD-UNIT-122 | unit | `tests/OfficeCli.Tests/Unit/WordNumberingAndLocaleTests.cs` | Word 编号格式边界、page defaults、locale font/RTL 默认值 |
| WORD-FREEZE-143 | WORD-UNIT-123/WORD-UNIT-124/WORD-UNIT-125 | unit | `tests/OfficeCli.Tests/Unit/WordTocBuilderTests.cs` | TOC heading/filter/bookmark/PAGEREF 基础重建已固定；refresh backend 生命周期、PDF page filter/grid/资源释放仍待补 |
| WORD-FREEZE-144 | WORD-UNIT-126/WORD-UNIT-127 | unit | `tests/OfficeCli.Tests/Unit/WordSchemaOrderTests.cs`, `tests/OfficeCli.Tests/Unit/WordDumpFilterTests.cs` | schema alias/order、synthetic property 过滤已部分固定，dump fold/warning 顺序仍待补 |
| WORD-FREEZE-145 | WORD-UNIT-128/WORD-UNIT-129 | unit/functional | `tests/OfficeCli.Tests/Functional/WordFieldViewContractTests.cs`, `tests/OfficeCli.Tests/Unit/WordRendererRegistryTests.cs` | view text 语义已有局部覆盖；内置 renderer capability、注册幂等、可用性 fallback 和错误 input 已固定 |
| WORD-FREEZE-146 | WORD-CMD-066/WORD-CMD-067 | e2e | `tests/OfficeCli.Tests/E2E/WordViewSmokeTests.cs` | mark/list/get-marks/unmark 当前显式路径、regex、颜色校验、version 和 `--all` 生命周期已固定；selection、SSE version/update、断开重连、完整错误 envelope 和 pipe 清理仍待补 |
| WORD-FREEZE-147 | WORD-CMD-068/WORD-CMD-069 | integration/e2e | `tests/OfficeCli.Tests/Integration/WordOpenRepairTests.cs`, `tests/OfficeCli.Tests/E2E/WordCreateOverwriteTests.cs` | open repair/corrupt 输入、create existing/force/type 冲突和文件不变性 |
| WORD-FREEZE-148 | WORD-CMD-070/WORD-CMD-071 | e2e | `tests/OfficeCli.Tests/E2E/WordViewBackendMatrixTests.cs`, `tests/OfficeCli.Tests/E2E/WordRefreshCliMatrixTests.cs` | view renderer/output 矩阵、refresh page/filter/backend/超时 envelope |
| WORD-FREEZE-149 | WORD-DOC-037/WORD-DOC-038 | contract | `tests/OfficeCli.Tests/Functional/WordViewSemanticsContractTests.cs`, `tests/OfficeCli.Tests/Functional/WordNumberingContractTests.cs` | 多 view mode 语义一致性、编号显示/继承/replay |
| WORD-FREEZE-150 | WORD-DOC-039/WORD-DOC-040 | contract | `tests/OfficeCli.Tests/Functional/WordSectionLinkageContractTests.cs`, `tests/OfficeCli.Tests/Functional/WordStyleNumberingInheritanceTests.cs` | section/header/footer/watermark 关系、style/numbering 继承和 ID 冲突 |
| WORD-FREEZE-151 | WORD-DOC-041/WORD-DOC-042 | contract/integration | `tests/OfficeCli.Tests/Functional/WordDrawingPropertyContractTests.cs`, `tests/OfficeCli.Tests/Integration/WordAuxiliaryDumpParityTests.cs` | drawing 参数清洗、auxiliary part 与 dump filter parity |
| WORD-FREEZE-152 | WORD-NEG-027/WORD-NEG-028 | integration/e2e | `tests/OfficeCli.Tests/Integration/WordOpenFailureAtomicityTests.cs`, `tests/OfficeCli.Tests/E2E/WordWatchMarkErrorTests.cs` | 打开修复/资源限制失败原子性、watch mark no-watch/reject 错误区分 |
| WORD-FREEZE-153 | WORD-NEG-029/WORD-NEG-030 | integration/e2e | `tests/OfficeCli.Tests/Integration/WordOutputPathBoundaryTests.cs`, `tests/OfficeCli.Tests/Integration/WordRelationshipRollbackTests.cs` | 输出/临时路径清理、relationship 冲突和 mutation 回滚 |
| WORD-FREEZE-154 | WORD-NEG-031 | integration | `tests/OfficeCli.Tests/Integration/WordResidentAtMostOnceTests.cs` | resident 半响应、超时、retry 和非幂等 mutation 不重复 |

## 首批建议落地顺序

| 顺序 | 范围 | 推荐挑选 |
|---:|---|---|
| 1 | P0 unit | `paragraph/run/picture/ole/sdt/selector` 的 handler/helper 直测 |
| 2 | P0 contract | `paragraph/run/table/cell/picture/sdt/revision/validate` |
| 3 | P0 integration | `create/add/set/get/query/remove/batch` 最小闭环 |
| 4 | P0 negative | 非法 path、非法 prop、缺 required prop、batch 失败策略 |
| 5 | 示例 smoke | `pictures/tables/content-controls/revisions` 各抽 1 条短链路 |
| 6 | round-trip regression | 图片、OLE、chart、header/footer rel 的 `dump -> batch` |
| 7 | E2E | 发布二进制执行 `create -> add -> view text -> validate -> close` |
| 8 | 视觉核验 | `view screenshot`、diagram、textbox、watermark、复杂 chart |
| 9 | command matrix | `create` 选项、`get --depth`、`view` 截断/分页/输出、batch file/stdin/坏输入 |
| 10 | field/document fidelity | TOC/PAGE/SEQ/REF refresh、未知 package part、section/header/footer 关联 |
| 11 | failure and locking | mutation 失败原子性、resident busy/stale、JSON stdout/stderr 纯度 |
| 12 | platform matrix | macOS/Windows 发布包、locale、CJK/RTL 内容与路径 |
