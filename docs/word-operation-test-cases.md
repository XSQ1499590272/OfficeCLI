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
| WORD-CMD-009 | integration | P0 | `README.md` | `validate` 检查生成后的 `.docx` | 当前新增用例覆盖多类生成后空错误；错误结构待补 | 已有自动覆盖 |
| WORD-CMD-010 | e2e | P1 | `README.md` | `view text/annotated/outline/stats/issues --json` | 真实 CLI 子进程输出 JSON 可解析；关键文本和各 mode 的稳定结构键存在 | 已有自动覆盖 |
| WORD-CMD-011 | e2e | P2 | `README.md` | `view html/screenshot` 渲染 Word 文档 | HTML 非空；截图文件存在且非空白 | 建议人工核验 |
| WORD-CMD-012 | integration | P1 | `README.md` | `open -> set -> save -> 外部读取 -> close` resident flush 链路 | `save` 后外部 OpenXML 读取到新值；`close` 后无锁文件 | 待补 |
| WORD-CMD-013 | integration | P1 | `README.md` | `OFFICECLI_RESIDENT_FLUSH=each` 每次修改落盘 | 每次 mutation 返回后磁盘内容已更新 | 待补 |
| WORD-CMD-014 | e2e | P2 | `README.md` | `watch` 本地预览服务 | 服务启动；HTML 端点可访问；修改后版本刷新 | 建议人工核验 |
| WORD-CMD-015 | e2e | P1 | `README.md` | `merge template.docx out.docx data.json` | 占位符跨段落、表格、页眉页脚被替换；未命中键保留或按约定处理 | 待补 |
| WORD-CMD-016 | integration | P1 | `schemas/help/docx/raw.json` | `raw` 读取 document/settings/styles 等 part | XPath 命中；输出 XML/JSON 可解析；不存在 part 有结构化报错 | 已有自动覆盖 |
| WORD-CMD-017 | integration | P1 | `schemas/help/docx/raw.json` | `raw-set` 修改单个 XML 节点或属性 | 写回后 `raw` 可读到新值；`validate` 通过；错误 XPath 不污染文件 | 已有自动覆盖 |
| WORD-CMD-018 | integration | P1 | `schemas/help/docx/raw.json` | `add-part` 创建新 part 和 relationship | chart part 可创建、raw 可读、main document relationship 指向 `ChartPart`；重复 chart part 生成不同 relId 和递增 path；CLI 路径待补 | 部分自动覆盖 |
| WORD-CMD-019 | e2e | P2 | `README.md` | `refresh` 更新 TOC / PAGE / cross-reference | 支持平台执行成功；不支持平台返回可解释降级 | 待补 |
| WORD-CMD-020 | e2e | P1 | `plugins/plugin-protocol.md` | 插件 `dump-reader` 产出 docx batch 并由主程序回放 | JSONL 流式输出；非法数组输出触发 `corrupt_batch` | 待补 |

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
| WORD-DOC-012 | contract | P2 | `schemas/help/docx/ptab.json` | header/footer 中 positional tab | body 段落 ptab add/set/get 已固定；header/footer 场景待补 | 已有自动覆盖 |
| WORD-DOC-013 | contract | P0 | `schemas/help/docx/style.json` | 新增 paragraph/character/table style | `get /styles/<id>` 返回类型、name、basedOn、格式属性；单元见 `tests/OfficeCli.Tests/Unit/WordStyleAndNumberingTests.cs` | 已有单元覆盖 |
| WORD-DOC-014 | contract | P1 | `schemas/help/docx/styles.json` | styles 容器查询和添加 style | `query style` 包含新增样式；重复 custom id 抛错且不污染 styles | 已有自动覆盖 |
| WORD-DOC-015 | integration | P1 | `examples/word/document-formatting.*` | 页面背景、默认字体、metadata 等文档格式示例 | 关键 document props 可读；`validate` 通过 | 示例待改造 |
| WORD-DOC-016 | contract | P0 | `schemas/help/docx/section.json` | section 尺寸、边距、方向、栏、页码、RTL gutter | `get /body/sectPr[1]` 返回 canonical 长度和方向值 | 已有自动覆盖 |
| WORD-DOC-017 | integration | P1 | `examples/word/sections.*` | 多 section 示例转集成测试 | section 数量、orientation、margin、columns 可读 | 示例待改造 |
| WORD-DOC-018 | contract | P1 | `schemas/help/docx/header.json` | 添加 default/first/even header | duplicate type 被拒绝；section 引用正确；内容可 query | 已有自动覆盖 |
| WORD-DOC-019 | contract | P1 | `schemas/help/docx/footer.json` | 添加 default/first/even footer | 与 header 语义一致；页脚内容可读写 | 已有自动覆盖 |

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
| WORD-MEDIA-001 | integration | P1 | `schemas/help/docx/picture.json` | CLI 查询 inline picture | handler 查询 inline picture 返回 `wrap=inline`；真实 CLI e2e 待补 | 已有单元覆盖 |
| WORD-MEDIA-002 | integration | P1 | `schemas/help/docx/picture.json` | CLI 查询 floating picture `wrap/hPosition/vPosition` | handler 查询 floating picture 返回 wrap 和相对定位；真实 CLI e2e 待补 | 已有单元覆盖 |
| WORD-MEDIA-003 | contract | P0 | `schemas/help/docx/picture.json` | add picture `src/width/height/alt` | `get` 返回尺寸和 alt；`validate` 通过 | 已有自动覆盖 |
| WORD-MEDIA-004 | contract | P1 | `schemas/help/docx/picture.json` | crop、decorative、link、behindText、wrap variants | wrap variants、floating position、crop、decorative、link、behindText 写入固定 | 已有自动覆盖 |
| WORD-MEDIA-005 | integration | P1 | `examples/word/pictures.*` | 图片示例转集成测试 | 8 个场景生成；关键图片路径和属性可读 | 示例待改造 |
| WORD-MEDIA-006 | integration | P1 | `schemas/help/docx/ole.json` | CLI 查询 OLE 对象并与 picture 区分 | handler 查询 OLE 返回 `progId/display/size`；真实 CLI e2e 待补 | 已有自动覆盖 |
| WORD-MEDIA-007 | contract | P1 | `schemas/help/docx/ole.json` | add OLE / embedded package | OLE rel、contentType、fileSize、尺寸可读；dump-batch 后 rel 不悬空待补 | 已有自动覆盖 |
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
| WORD-MEDIA-019 | contract | P1 | `schemas/help/docx/watermark.json` | text/image watermark | text watermark add/get/query/set/remove、VML 属性读回、非法 rotation 固定；handler 层 `image` 当前退回默认 text watermark 的行为已固定，真实 image watermark 待补；单元见 `tests/OfficeCli.Tests/Unit/WordComplexObjectTests.cs` | 已有单元覆盖 |
| WORD-MEDIA-020 | regression | P0 | `README.md` | `dump -> batch` 保留图片 part 和 rel | 新文档图片可 `query/get`、image part bytes 与源 data URI 一致，且 `validate` 通过 | 已有自动覆盖 |
| WORD-MEDIA-021 | regression | P0 | `README.md` | `dump -> batch` 保留 OLE / embedded package | `query ole` 返回 progId/name/contentType/fileSize；`validate` 通过 | 已有自动覆盖 |
| WORD-MEDIA-022 | regression | P1 | `README.md` | `dump -> batch` 保留 chart 的 externalData / embedded workbook | chart/series/cache 可读且 `validate` 通过；workbook rel 存在待补 | 部分自动覆盖 |
| WORD-MEDIA-023 | regression | P1 | `README.md` | header/footer 内图片或 textbox round-trip | header 内图片 dump/batch 后可在 header 子树读回 name/alt；footer/textbox 待补 | 部分自动覆盖 |

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

## 修订与回归场景

| ID | 层级 | 优先级 | 依据 | 操作 / 场景 | 建议断言 | 状态 |
|---|---|---:|---|---|---|---|
| WORD-REV-001 | contract | P0 | `schemas/help/docx/revision.json` | run insertion/deletion/format revision | `query revision` 返回 type、author、id、path、text；date 字段形态待补 | 已有自动覆盖 |
| WORD-REV-002 | contract | P0 | `schemas/help/docx/revision.json` | paragraph add/remove/set with `revision.author` | paragraph format revision 可查询；synthetic revision 当前读回 `revision.type=paragraph`，host paragraph 读回 `format`；当前 paragraph `revision.type=ins` via set 明确拒绝；单元见 `tests/OfficeCli.Tests/Unit/WordRevisionTests.cs`；add/remove tracked paragraph 待补 | 已有单元覆盖 |
| WORD-REV-003 | contract | P1 | `schemas/help/docx/revision.json` | moveFrom/moveTo 配对 `revision.id` | run 级 moveFrom/moveTo 两半共享 id；synthetic revision path 按 type 消歧；单元见 `tests/OfficeCli.Tests/Unit/WordRevisionTests.cs`；range marker 细节待补 | 已有单元覆盖 |
| WORD-REV-004 | contract | P1 | `schemas/help/docx/revision.json` | table/row/cell scope revision | table/cell format marker、row 级 ins marker 可查询，synthetic row type 当前为 `rowIns`、host row 为 `ins`；单元见 `tests/OfficeCli.Tests/Unit/WordRevisionTests.cs`；trPr format 和 cell del/ins 待补 | 已有单元覆盖 |
| WORD-REV-005 | contract | P1 | `schemas/help/docx/revision.json` | section property revision | `sectPrChange` author/id 可读；section orientation 更新可读；synthetic revision type 当前为 `format` | 已有自动覆盖 |
| WORD-REV-006 | contract | P1 | `schemas/help/docx/revision.json` | accept/reject 单个 revision | run insertion/deletion/format/move accept/reject 当前结果固定 | 已有自动覆盖 |
| WORD-REV-007 | integration | P1 | `examples/word/revisions.*` | revisions 示例转集成测试 | 8 个 revision 场景生成；accept/reject temp copy 验证 | 示例待改造 |
| WORD-REV-008 | regression | P0 | `tests/OfficeCli.Tests/UnitTest1.cs` | picture 与 OLE query 不互相污染 | picture 数量和 ole 数量稳定 | 已有单元覆盖 |
| WORD-REV-009 | regression | P0 | `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs` | SDT schema/runtime 漂移防护 | schema 声明和 handler readback 不漂移 | 已有单元覆盖 |
| WORD-REV-010 | regression | P1 | `README.md` | schema 中 canonical key 与 legacy alias | shape set 的 `preset/fillcolor/linecolor/linewidth` legacy alias 写入后 raw canonical 输出稳定；其他 alias 待补 | 部分自动覆盖 |
| WORD-REV-011 | regression | P1 | `README.md` | 错误路径、非法属性、非法枚举值 | chart series 非法数值抛 `invalid_value` 但当前会清空 values 并留下 schema 错误；其他场景待补 | 已有自动覆盖 |

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
