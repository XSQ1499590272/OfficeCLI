# OfficeCLI agent-json 阶段 0 交付证据

Author: xiesq
Updated: 2026-07-16

## 身份

| 项 | 值 |
| --- | --- |
| 仓库 | `office-cli/agnet-officecli` |
| 分支 | `develop` |
| Agent 功能 commit | `b6de11c85b81935aa0a395ab9de30a3219dbb527` |
| 独立 Excel 测试调整 | `fd8b522` |
| 独立资源型超时调整 | `4daa95e` |
| 包版本 | `1.0.132` |
| Surface | `--surface agent-json` |
| 角色占位符 | `{{OFFICE_HELP_TOOL}}` / `{{OFFICE_LOAD_SKILL_TOOL}}` / `{{OFFICE_BATCH_TOOL}}` / `{{OFFICE_RUN_TOOL}}` |

## 实现与源 Skill 一致性

- 10 个 `SkillMap` 条目均有独立 `agent-skills/` 资源和 manifest。
- L3 已逐一对比源 `skills/`：业务规则、限制、已知陷阱语义一致；把 CLI/Shell recipe 改写为宿主 Tool JSON，以及增加阶段 0 安全边界，属于约定内差异。
- `skills/officecli/` 不属于 `SkillMap`，因此没有 Agent 对应项，符合当前计划范围。
- Morph style id：源目录与 Agent 目录 `54/54` 完全对应。
- 373 个 JSON fence 通过精确 envelope 校验：Help 29、LoadSkill 7、Run 298、Batch 39。
- 10 个 manifest 的固定结构、名称、reference allowlist、资源存在性与 128 KiB 文本上限检查通过。
- Batch 独立性、生成 ID 依赖、重叠写入、protection 隔离，以及禁止 CLI/Shell recipe、prose option、未知 Tool token 的检查通过。

## 测试证据

| 验证 | 结果 |
| --- | --- |
| Agent guidance 聚焦测试 | `80/80` 通过 |
| `OfficeCli.Main.Tests` | `17/17` 通过 |
| `OfficeCli.Tests` 全量 | `1656/1656` 通过，0 跳过，耗时 8 分 42 秒 |
| `WordPublishedBinarySmokeTests` | `1/1` 通过，耗时 8 秒 |
| Word refresh 定向复测 | `1/1` 通过，耗时 36 秒 |
| staged diff check | 通过 |
| L3 双人静态审查 | 通过，无新增 blocker |

全量命令：

```text
/usr/local/share/dotnet/dotnet test officecli.slnx --nologo --no-restore
```

发布后二进制 smoke 在 macOS 主机上交叉 publish `osx-arm64` 和 `win-x64`，并原生执行 macOS 产物的 Help/catalog/core skill/Morph reference 检查。macOS Mach-O portability gate 通过。

## 同一 commit/version 的本地产物

以下产物均从 detached clean worktree 的 `b6de11c85b81935aa0a395ab9de30a3219dbb527`、版本 `1.0.132` 构建：

| 产物 | 字节数 | SHA-256 |
| --- | ---: | --- |
| `officecli-1.0.132-osx-arm64.tar.gz` | 12,878,082 | `5d327673484e7cf5008e41c7d9c9228f0499f1134c60dd2c1a172820535ab1cc` |
| `officecli-1.0.132-win-x64.zip` | 13,178,285 | `4249834c3c1e1e7ec7ccdc80eb93ef4f9c0819caf711663795447a4232223577` |

本地临时目录：`/tmp/officecli-agent-b6de11c-artifacts/`。这些哈希证明同 commit/version 的可交叉构建性，不替代 tag workflow 的正式发布 artifact。

## Agent Skill 行数（当前 commit）

| Skill | 行数 |
| --- | ---: |
| `morph-ppt-3d` | 547 |
| `morph-ppt` | 814 |
| `officecli-academic-paper` | 828 |
| `officecli-data-dashboard` | 598 |
| `officecli-docx` | 1323 |
| `officecli-financial-model` | 707 |
| `officecli-pitch-deck` | 2461 |
| `officecli-pptx` | 1268 |
| `officecli-word-form` | 1592 |
| `officecli-xlsx` | 887 |

`morph-ppt` references：decision 146、design 247、style catalog 134 行。

## 尚未完成的正式发布门禁

当前仓库只配置 Gitee remote，没有可触发 `.github/workflows/build.yml` 的 GitHub remote，也没有可用的 Windows x64 原生 runner。因此：

1. 尚未打正式发布 tag；
2. Windows x64 已交叉 publish，但尚未原生执行 smoke；
3. 上述本地 SHA-256 尚不是 tag workflow 发布页上的正式 artifact SHA。

在 GitHub workflow 和 Windows runner 可用前，阶段 0 实现与验证可以提交，但不得把正式发布/O05 标记为完成。

## 证据有效性说明

`docs/delivery/` 中 2026-07-15 的旧日志和旧 package 不是本次稳定 commit 的正式证据，且部分旧文案引用了不存在的日志；它们未纳入提交。本文件取代其中过期的 `75/75`、旧行数和未提交工作树描述。
