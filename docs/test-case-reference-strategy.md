# 测试用例参考策略

本文用于指导后续编写测试用例清单，目标是避免只通过源码倒推测试，而是从项目已经公开承诺的能力、示例流程、协议约束和 CI 验证路径中抽取测试依据。

## 目标

- 建立测试用例的来源优先级。
- 明确单元测试、集成测试、自动 E2E 测试各自应该覆盖什么。
- 为后续测试用例清单提供统一的记录格式。

## 测试依据优先级

按下面顺序寻找测试依据。源码用于确认实现边界和补充细节，但不应作为唯一依据。

1. `schemas/help/**/*.json`
   - 这是 agent-facing 能力的事实来源。
   - 适合生成契约测试：schema 宣称支持的 `add` / `set` / `get` / `readback` 能力，应能通过真实 handler 或 CLI 命令验证。

2. `examples/**/*.md`、`examples/**/*.sh`、`examples/**/*.py`
   - 这是最接近真实用户流程的材料。
   - 适合转成集成测试和轻量 E2E：执行示例脚本，随后用 `validate`、`get`、`query`、`view` 做关键断言。

3. `.github/workflows/*.yml`
   - 这是当前 CI 实际执行的验证边界。
   - 适合定义 E2E 的最低线：发布产物能创建文档、修改文档、读取文档、安装自身，SDK 能完成自动安装和 pipe round-trip。

4. `README.md`、`plugins/plugin-protocol.md`、`schemas/README.md`
   - 这是架构、协议和能力说明。
   - 适合提取跨模块测试：L1 view、L2 DOM、L3 raw XML、插件协议、错误码、session 生命周期。

5. 历史回归信息
   - 历史提交中带有 `fix(...)`、`round-trip`、`dump-batch test campaign`、`schema drift` 等关键词的提交，可作为回归测试候选来源。
   - 只把稳定可复现的问题纳入自动测试；一次性排查记录不要原样搬进测试套件。

6. 源码
   - 用于确认具体输入、边界条件、异常类型、格式化细节。
   - 不把私有实现细节当成外部契约，除非该行为已经被 schema、README、示例或协议文档承诺。

## 测试分层

### 单元测试

适合直接调用 handler、helper 或纯逻辑类，不启动 `officecli` 子进程。

优先覆盖：

- 解析、格式化、单位换算、路径选择器、schema key 规范化。
- Word / Excel / PowerPoint handler 中可用临时 OOXML 文档构造的局部行为。
- 历史上容易漂移的 readback key、alias、错误分支。

参考现有文件：

- `tests/OfficeCli.Tests/UnitTest1.cs`
- `tests/OfficeCli.Tests/Functional/WordSdtSchemaHonestyTests.cs`

### 集成测试

适合通过 `officecli` 命令或 command root 执行完整读写链路，但不要求使用发布后的真实安装包。

优先覆盖：

- `create -> add/set -> get/query -> validate`。
- `batch` 与单命令行为一致性。
- `dump -> batch -> validate` 的小型 round-trip。
- `schemas/help` 中声明的关键 add/set/get/readback 能力。

推荐从 `examples/**/*.sh` 和 `examples/**/*.py` 中抽取短链路，不直接全量照搬大型示例。

### 自动 E2E 测试

适合使用发布或构建后的真实二进制，从用户视角验证端到端行为。

优先覆盖：

- 三种主格式各一条最小链路：`.docx`、`.xlsx`、`.pptx`。
- 安装链路：`install`、`install.sh`、`install.ps1`。
- SDK 链路：Node SDK 和 Python SDK 的自动安装 + pipe round-trip。
- 少量代表性示例脚本，而不是全量 examples。

参考现有文件：

- `.github/workflows/build.yml`
- `.github/workflows/sdk-smoke.yml`
- `sdk/node/smoke.js`
- `sdk/python/smoke.py`

## 用例清单记录格式

后续编写测试用例清单时，每条用例建议至少记录：

```markdown
### TC-<format>-<module>-<number>: <一句话目标>

- 类型: unit | integration | e2e | contract | regression
- 优先级: P0 | P1 | P2
- 依据: schema | example | ci | readme | plugin-protocol | regression | source-confirmed
- 依据文件: `<path>`
- 覆盖能力: `<command/element/property>`
- 前置条件: `<需要的样例文件、SDK、平台或环境变量>`
- 步骤:
  1. `<step>`
  2. `<step>`
- 断言:
  - `<assertion>`
- 不覆盖:
  - `<明确排除的重型或不稳定场景>`
```

## 首批补齐顺序

1. 先补 schema contract smoke
   - 每个格式挑 3 到 5 个高频元素。
   - 先验证 schema 声明和 handler readback 不漂移。

2. 再补 examples integration smoke
   - 每个格式挑 2 个代表示例。
   - 只断言关键输出，不做像素级或完整 XML 对比。

3. 再补发布二进制 E2E
   - 复用现有 CI smoke。
   - 增加最小 `.xlsx` 和 `.pptx` 链路，避免只测 `.docx`。

4. 最后补历史回归
   - 从最近高风险修复中挑稳定、短小、可复现的案例。
   - 每个回归测试应能说明防止哪个已知行为再次坏掉。

## 不建议做的事

- 不要直接把全部 `examples/` 都接进默认 CI，成本和不稳定性会很快失控。
- 不要只因为源码里有分支就写测试；先确认它是不是外部契约。
- 不要对完整 OOXML 做大段字符串断言；优先用 `get`、`query`、`validate`、OpenXML SDK 结构化读取。
- 不要把视觉效果默认做成自动阻塞测试；先用 `view screenshot` 或人工检查作为高层验证，只有稳定像素输出再考虑自动化。
