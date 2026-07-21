// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Author: xiesq
// Created: 2026-07-15

using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using OfficeCli.Help;

namespace OfficeCli.Tests.Unit;

/// <summary>
/// Phase-0 Agent guidance conformance gate (O05).
/// Exact-shape fixture checker — not a production canonical contract.
/// Author: xiesq, 2026-07-15
/// </summary>
[Trait("Speed", "Unit")]
public sealed class AgentGuidanceConformanceTests
{
    private static readonly string[] RoleTokens =
    [
        GuidanceSurface.HelpTool,
        GuidanceSurface.LoadSkillTool,
        GuidanceSurface.BatchTool,
        GuidanceSurface.RunTool,
    ];

    private static readonly HashSet<string> BatchMutationCommands =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "set", "add", "import", "remove", "move", "swap", "raw-set", "add-part",
        };

    private static readonly HashSet<string> ReadHeavyCommands =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "get", "query", "view", "raw", "validate",
        };

    private static readonly string[] AgentHelpVerbs =
        ["add", "set", "get", "query", "remove", "import", "move", "swap", "view", "raw", "validate"];

    private static readonly HashSet<string> AllowedOperationFields =
        new(StringComparer.Ordinal)
        {
            "command", "path", "parent", "type", "from", "index", "after", "before",
            "to", "path2", "props", "text", "part", "xpath", "action", "xml",
        };

    private static readonly Regex JsonFenceRegex =
        new(@"```json\s*(\{[\s\S]*?\})\s*```", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex OfficeCliCommandRegex =
        new(@"\bofficecli\s+\S+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex UnknownOfficeTokenRegex =
        new(@"\{\{OFFICE_[A-Z0-9_]+\}\}", RegexOptions.Compiled);

    private static readonly Regex CliOptionTokenRegex =
        new(@"(?<![A-Za-z0-9_])--[A-Za-z][A-Za-z0-9-]*", RegexOptions.Compiled);

    private static readonly Regex ShellEscapeRegex = new(
        @"(?:\b(?:bash|sh|zsh)\s+-c\b"
        + @"|\b(?:powershell|pwsh)\b[^\r\n]{0,48}\s-(?:command|file)\b"
        + @"|\bcmd(?:\.exe)?\s+/c\b"
        + @"|\bProcess\.Start\b|\bsubprocess\.(?:run|Popen)\b|\bos\.system\b"
        + @"|(?:\./|\b(?:python3?|bash|sh)\s+)\S+\.(?:sh|py|ps1)\b"
        + @"|(?:\s2>|\s>>|\s>\s*(?:/|[A-Za-z0-9_.-]+\.(?:txt|log|json|xml|html|csv|md)\b))[^\r\n]*)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly string[] ForbiddenPatterns =
    [
        "--commands",
        "$FILE",
        "tool.file.office.",
        "```bash",
        "```sh",
        "```shell",
        "```python",
        "```powershell",
        "| jq",
        "| grep",
        "<<EOF",
    ];

    [Fact]
    public void AgentGuidance_AllHelpSchemasConform()
    {
        foreach (var format in SchemaHelpLoader.ListFormats())
        {
            foreach (var element in SchemaHelpLoader.ListElements(format))
            {
                using var schema = SchemaHelpLoader.LoadSchema(format, element);
                var text = SchemaHelpAgentRenderer.RenderElement(schema);
                AssertGuidanceText($"help:{format}/{element}", text, requireRoutingSemantics: false);
                foreach (var verb in AgentHelpVerbs)
                {
                    var verbText = SchemaHelpAgentRenderer.RenderElement(schema, verb);
                    AssertGuidanceText(
                        $"help:{format}/{element}?verb={verb}",
                        verbText,
                        requireRoutingSemantics: false);
                }
            }
        }

        AssertGuidanceText("help:banner", SchemaHelpAgentRenderer.RenderBanner(), requireRoutingSemantics: true);
    }

    [Fact]
    public void AgentGuidance_AgentCatalogEqualsDefaultSkillNames()
    {
        var expected = SkillCatalog.SkillMapNamesInOrder().ToList();
        var actual = SkillCatalog.ListAgentSkillNames().ToList();
        Require(expected.Count == actual.Count, "catalog-count", "catalog");
        Require(expected.SequenceEqual(actual), "catalog-order", "catalog");
        Require(
            actual.Count == actual.Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            "catalog-duplicate", "catalog");
    }

    [Fact]
    public void AgentGuidance_AllSkillsAndReferencesConform()
    {
        foreach (var name in SkillCatalog.ListAgentSkillNames())
        {
            var main = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            AssertGuidanceText($"skill:{name}", main, requireRoutingSemantics: true);

            var refs = SkillCatalog.ListSkillFiles(name, GuidanceSurfaceKind.AgentJson);
            foreach (var path in refs)
            {
                var content = SkillCatalog.LoadSkillFile(name, path, GuidanceSurfaceKind.AgentJson);
                AssertGuidanceText($"skill:{name}/{path}", content, requireRoutingSemantics: false);
            }
        }
    }

    [Fact]
    public void AgentGuidance_AllToolEnvelopesUseExactShapes()
    {
        foreach (var (id, json) in EnumerateAllAgentJsonFences())
            AssertExactToolEnvelope(id, json);
    }

    [Fact]
    public void AgentGuidance_BatchExamplesUseMutationFastPath()
    {
        foreach (var (id, json) in EnumerateAllAgentJsonFences())
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.BatchTool)
                continue;
            var ops = doc.RootElement.GetProperty("arguments").GetProperty("operations");
            Require(ops.GetArrayLength() >= 1, "empty-operations", id);
            // guidance heuristic：示例通常 ≥3；允许重试场景短数组不在本门禁硬拒绝
            foreach (var op in ops.EnumerateArray())
                AssertBatchOperation(id, op);
            AssertBatchOperationsIndependent(id, ops);
        }
    }

    [Fact]
    public void AgentGuidance_BatchExamplesEnforceCrossFieldRules()
    {
        using var ok = JsonDocument.Parse("""{"command":"add","parent":"/body","type":"paragraph","index":1}""");
        Require(!SchemaHelpAgentRenderer.HasAmbiguousPositionFields(ok.RootElement), "fixture-ok-position", "fixture:ok");
        Require(!SchemaHelpAgentRenderer.HasAddFromAndProps(ok.RootElement), "fixture-ok-from-props", "fixture:ok");

        using var ambiguous = JsonDocument.Parse(
            """{"command":"move","path":"/a","index":1,"after":"/b"}""");
        Require(SchemaHelpAgentRenderer.HasAmbiguousPositionFields(ambiguous.RootElement), "fixture-missed-position", "fixture:ambiguous");

        using var fromProps = JsonDocument.Parse(
            """{"command":"add","parent":"/body","type":"picture","from":"/a.png","props":{"w":"1"}}""");
        Require(SchemaHelpAgentRenderer.HasAddFromAndProps(fromProps.RootElement), "fixture-missed-from-props", "fixture:from-props");
    }

    [Fact]
    public void AgentGuidance_RoutingUsesRunForSingleReadAndDiagnosticSteps()
    {
        foreach (var (id, json) in EnumerateAllAgentJsonFences())
        {
            using var doc = JsonDocument.Parse(json);
            var tool = doc.RootElement.GetProperty("tool").GetString();
            var args = doc.RootElement.GetProperty("arguments");
            if (tool == GuidanceSurface.BatchTool)
            {
                var ops = args.GetProperty("operations");
                // Skill/banner 示例应通常 ≥3；Help 单命令页不应含 Batch
                if (id.StartsWith("help:", StringComparison.Ordinal)
                    && !id.Equals("help:banner", StringComparison.Ordinal))
                    throw RuleError("help-emits-batch", id);
                if (id.StartsWith("skill:", StringComparison.Ordinal)
                    || id.Equals("help:banner", StringComparison.Ordinal))
                    Require(ops.GetArrayLength() >= 3, "batch-too-short", id);
            }
            else if (tool == GuidanceSurface.RunTool)
            {
                var cmd = args.GetProperty("command_name").GetString()!;
                Require(!cmd.Equals("batch", StringComparison.OrdinalIgnoreCase), "run-batch", id);
            }
        }
    }

    [Fact]
    public void AgentGuidance_RunExamplesRejectOnlyBatchCommand()
    {
        var allowedSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (id, json) in EnumerateAllAgentJsonFences())
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.RunTool)
                continue;
            var cmd = doc.RootElement.GetProperty("arguments").GetProperty("command_name").GetString()!;
            Require(!cmd.Equals("batch", StringComparison.OrdinalIgnoreCase), "run-batch", id);
            allowedSeen.Add(cmd);
        }

        Require(allowedSeen.Contains("validate"), "missing-run-validate", "corpus");
        Require(allowedSeen.Contains("get"), "missing-run-get", "corpus");
    }

    [Fact]
    public void AgentGuidance_HighRiskFallbacksRequireApprovalAndNoOutputDependency()
    {
        var corpus = CollectAllAgentText();
        Require(corpus.Contains("approval", StringComparison.OrdinalIgnoreCase), "missing-approval", "corpus");
        Require(corpus.Contains("raw-set", StringComparison.OrdinalIgnoreCase), "missing-raw-set", "corpus");
        Require(corpus.Contains("add-part", StringComparison.OrdinalIgnoreCase), "missing-add-part", "corpus");
        // 结果依赖后续步骤必须拆 Run
        Require(corpus.Contains("relationship", StringComparison.OrdinalIgnoreCase), "missing-relationship", "corpus");

        // raw-set/add-part 可出现在 allowlist/prose 风险说明中，但不得作为普通 Batch 首选 recipe（JSON fence）
        foreach (var (id, json) in EnumerateAllAgentJsonFences())
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.BatchTool)
                continue;

            foreach (var op in doc.RootElement.GetProperty("arguments").GetProperty("operations").EnumerateArray())
            {
                var cmd = op.GetProperty("command").GetString()!;
                Require(
                    !cmd.Equals("raw-set", StringComparison.OrdinalIgnoreCase)
                    && !cmd.Equals("add-part", StringComparison.OrdinalIgnoreCase),
                    "high-risk-batch", id, $"cmd={cmd}");
            }
        }
    }

    [Fact]
    public void AgentGuidance_RequiresPartialSuccessBoundaryAndIndependentValidation()
    {
        Require(SchemaHelpAgentRenderer.RenderBanner().Contains("不是事务", StringComparison.Ordinal), "missing-nontransaction", "help:banner");
        foreach (var name in SkillCatalog.ListAgentSkillNames())
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            var id = $"skill:{name}";
            Require(text.Contains("不是事务", StringComparison.Ordinal), "missing-nontransaction", id);
            Require(text.Contains("不回滚", StringComparison.Ordinal), "missing-no-rollback", id);
            Require(text.Contains("discardable", StringComparison.OrdinalIgnoreCase), "missing-discardable", id);
            Require(text.Contains("validate", StringComparison.OrdinalIgnoreCase), "missing-validation", id);
            Require(text.Contains("不会自动创建草稿", StringComparison.Ordinal), "missing-draft-boundary", id);
        }
    }

    [Fact]
    public void BatchHarness_StopOnErrorDoesNotRollbackPriorSuccess()
    {
        // 临时副本上验证：成功项保留、失败后后续不执行；不宣称生产草稿生命周期
        var path = Path.Combine(Path.GetTempPath(), $"officecli_batch_soe_{Guid.NewGuid():N}.docx");
        try
        {
            BlankDocCreator.Create(path);
            const string itemsJson = """
            [
              {"command":"add","parent":"/body","type":"paragraph","props":{"text":"kept-success"}},
              {"command":"add","parent":"/body/p[99]","type":"paragraph","props":{"text":"fail"}},
              {"command":"add","parent":"/body","type":"paragraph","props":{"text":"should-skip"}}
            ]
            """;

            using var handler = new WordHandler(path, editable: true);
            var output = BatchExecutor.ExecuteBatch(handler, itemsJson, json: false, stopOnError: true);

            Require(output.Contains("ERROR", StringComparison.OrdinalIgnoreCase), "batch-harness-no-error", "batch-harness");
            Require(handler.Query("paragraph:contains(\"kept-success\")").Count == 1, "batch-harness-rollback", "batch-harness");
            Require(handler.Query("paragraph:contains(\"should-skip\")").Count == 0, "batch-harness-did-not-stop", "batch-harness");
        }
        finally
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void AgentGuidance_ReadHeavyBatchAndOverflowAreExplicitlyExcluded()
    {
        foreach (var (id, json) in EnumerateAllAgentJsonFences())
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.BatchTool)
                continue;
            foreach (var op in doc.RootElement.GetProperty("arguments").GetProperty("operations").EnumerateArray())
            {
                var cmd = op.GetProperty("command").GetString()!;
                Require(!ReadHeavyCommands.Contains(cmd), "read-heavy-batch", id, $"cmd={cmd}");
            }
        }

        var banner = SchemaHelpAgentRenderer.RenderBanner();
        Require(banner.Contains("8192", StringComparison.Ordinal), "missing-overflow-limit", "help:banner");
        Require(banner.Contains("outputFile", StringComparison.Ordinal), "missing-output-file", "help:banner");
    }

    [Fact]
    public void ReadmeBatchSemantics_MatchCurrentImplementation()
    {
        var readme = File.ReadAllText(Path.Combine(FindRepoRoot(), "README_zh.md"));
        foreach (var stale in new[]
                 {
                     "原子化多命令执行", "默认遇到第一个错误即停止", "默认遇到第一个错误停止",
                     "--force 跳过错误继续", "`--force` 跳过错误继续",
                 })
            Require(!readme.Contains(stale, StringComparison.Ordinal), "readme-stale-semantics", "README_zh.md", $"value={stale}");

        foreach (var required in new[] { "默认遇错继续", "--stop-on-error", "不回滚", "文档保护" })
            Require(readme.Contains(required, StringComparison.Ordinal), "readme-missing-semantics", "README_zh.md", $"value={required}");
    }

    [Fact]
    public void AgentGuidance_UsesOnlyKnownRolePlaceholders()
    {
        var text = CollectAllAgentText();
        foreach (Match m in UnknownOfficeTokenRegex.Matches(text))
            Require(RoleTokens.Contains(m.Value), "unknown-tool-token", "corpus", $"token={m.Value}");
        Require(!text.Contains("tool.file.office.", StringComparison.Ordinal), "concrete-tool-name", "corpus");
    }

    [Fact]
    public void AgentGuidance_NegativeFixturesTriggerEveryRule()
    {
        AssertRule("envelope-exact-shape", () => AssertExactToolEnvelope("neg",
            """{"tool":"{{OFFICE_HELP_TOOL}}","arguments":{"command_arguments":[]},"extra":true}"""));

        AssertRule("unknown-tool", () => AssertExactToolEnvelope("neg",
            """{"tool":"{{OFFICE_UNKNOWN_TOOL}}","arguments":{}}"""));

        AssertRule("commands-string", () => AssertExactToolEnvelope("neg",
            """{"tool":"{{OFFICE_BATCH_TOOL}}","arguments":{"file":"/a.docx","operations":[],"stop_on_error":true,"commands":"[]"}}"""));

        AssertRule("empty-operations", () => AssertExactToolEnvelope("neg",
            """{"tool":"{{OFFICE_BATCH_TOOL}}","arguments":{"file":"/a.docx","operations":[],"stop_on_error":true}}"""));

        // rule:props-null
        using var nullProps = JsonDocument.Parse(
            """{"command":"set","path":"/a","props":null}""");
        AssertRule("props-null", () => AssertBatchOperation("neg", nullProps.RootElement));

        // rule:read-in-batch
        using var readOp = JsonDocument.Parse("""{"command":"get","path":"/a"}""");
        AssertRule("read-heavy-batch", () => AssertBatchOperation("neg", readOp.RootElement));

        using var unknownField = JsonDocument.Parse(
            """{"command":"set","path":"/a","unknown":true}""");
        AssertRule("unknown-field", () => AssertBatchOperation("neg", unknownField.RootElement));

        using var opAlias = JsonDocument.Parse(
            """{"op":"set","path":"/a","props":{"bold":true}}""");
        AssertRule("op-alias", () => AssertBatchOperation("neg", opAlias.RootElement));

        using var compoundProps = JsonDocument.Parse(
            """{"command":"set","path":"/a","props":{"nested":{"value":1}}}""");
        AssertRule("props-compound", () => AssertBatchOperation("neg", compoundProps.RootElement));

        using var ambiguousPosition = JsonDocument.Parse(
            """{"command":"move","path":"/a","index":1,"after":"/b"}""");
        AssertRule(
            "ambiguous-position", () => AssertBatchOperation("neg", ambiguousPosition.RootElement));

        using var addFromProps = JsonDocument.Parse(
            """{"command":"add","parent":"/body","type":"picture","from":"/a.png","props":{"width":"1in"}}""");
        AssertRule("add-from-props", () => AssertBatchOperation("neg", addFromProps.RootElement));

        // rule:run-batch
        AssertRule("run-batch", () => AssertExactToolEnvelope("neg",
            """{"tool":"{{OFFICE_RUN_TOOL}}","arguments":{"command_name":"batch","command_arguments":[]}}"""));

        // rule:option-in-prose
        AssertRule("option-in-prose", () => AssertGuidanceText(
            "neg", "Run set / --prop value=true", requireRoutingSemantics: false));

        foreach (var escape in new[]
                 {
                     "bash -c do_work",
                     "python script.py",
                     "pwsh -File run.ps1",
                     "Process.Start(executable)",
                     "subprocess.run(command)",
                     "result > /tmp/output.txt",
                     "result > output.txt",
                 })
            AssertRule("shell-escape", () => AssertGuidanceText(
                "neg", escape, requireRoutingSemantics: false));

        // rule:batch-dependency
        using var dependent = JsonDocument.Parse("""
            [
              {"command":"add","parent":"/body","type":"table","props":{"rows":2}},
              {"command":"set","path":"/body/tbl[1]/tr[1]","props":{"header":true}}
            ]
            """);
        AssertRule("batch-dependency", () =>
            AssertBatchOperationsIndependent("neg", dependent.RootElement));

        using var overlapping = JsonDocument.Parse("""
            [
              {"command":"set","path":"/body/sdt[1]","props":{"lock":"sdtlocked"}},
              {"command":"set","path":"/body/sdt[1]","props":{"lock":"contentlocked"}}
            ]
            """);
        AssertRule("batch-overlapping-write", () =>
            AssertBatchOperationsIndependent("neg", overlapping.RootElement));

        using var generatedId = JsonDocument.Parse("""
            [
              {"command":"add","parent":"/numbering","type":"abstractnum","props":{"format":"decimal"}},
              {"command":"add","parent":"/numbering","type":"num","props":{"abstractNumId":0}}
            ]
            """);
        AssertRule("batch-generated-id-dependency", () =>
            AssertBatchOperationsIndependent("neg", generatedId.RootElement));
    }

    [Fact]
    public void AgentGuidance_LargeCorpusCompletesWithinBudget()
    {
        var sw = Stopwatch.StartNew();
        for (var i = 0; i < 2; i++)
        {
            _ = SchemaHelpAgentRenderer.RenderBanner();
            foreach (var format in SchemaHelpLoader.ListFormats())
            foreach (var element in SchemaHelpLoader.ListElements(format))
            {
                using var schema = SchemaHelpLoader.LoadSchema(format, element);
                _ = SchemaHelpAgentRenderer.RenderElement(schema);
            }

            foreach (var name in SkillCatalog.ListAgentSkillNames())
            {
                _ = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
                foreach (var path in SkillCatalog.ListSkillFiles(name, GuidanceSurfaceKind.AgentJson))
                    _ = SkillCatalog.LoadSkillFile(name, path, GuidanceSurfaceKind.AgentJson);
            }
        }

        sw.Stop();
        Require(
            sw.Elapsed < TimeSpan.FromSeconds(30),
            "performance-budget", "corpus", $"elapsed={sw.Elapsed}");
    }

    [Fact]
    public void AgentGuidanceWorkflow_TriggersForEveryOwnedPath()
    {
        var workflow = File.ReadAllText(
            Path.Combine(FindRepoRoot(), ".github", "workflows", "agent-guidance.yml"));
        foreach (var path in new[]
                 {
                     "README_zh.md",
                     "src/officecli/**",
                     "schemas/help/**",
                     "skills/**",
                     "agent-skills/**",
                     "tests/OfficeCli.Tests/**",
                     ".github/workflows/**",
                 })
            Require(workflow.Contains(path, StringComparison.Ordinal), "workflow-missing-path", "agent-guidance.yml", $"path={path}");

        Require(
            !workflow.Contains("continue-on-error: true", StringComparison.Ordinal),
            "workflow-continue-on-error", "agent-guidance.yml");
    }

    // ── helpers ──────────────────────────────────────────────────────────

    private static IEnumerable<(string Id, string Text)> EnumerateAllAgentTexts()
    {
        yield return ("help:banner", SchemaHelpAgentRenderer.RenderBanner());
        foreach (var format in SchemaHelpLoader.ListFormats())
        {
            foreach (var element in SchemaHelpLoader.ListElements(format))
            {
                using var schema = SchemaHelpLoader.LoadSchema(format, element);
                yield return ($"help:{format}/{element}", SchemaHelpAgentRenderer.RenderElement(schema));
            }
        }

        foreach (var name in SkillCatalog.ListAgentSkillNames())
        {
            yield return ($"skill:{name}", SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson));
            foreach (var path in SkillCatalog.ListSkillFiles(name, GuidanceSurfaceKind.AgentJson))
                yield return (
                    $"skill:{name}/{path}",
                    SkillCatalog.LoadSkillFile(name, path, GuidanceSurfaceKind.AgentJson));
        }
    }

    private static IEnumerable<(string Id, string Json)> EnumerateAllAgentJsonFences()
    {
        foreach (var (id, text) in EnumerateAllAgentTexts())
        {
            foreach (var fence in ExtractJsonFences(text))
                yield return (id, fence);
        }
    }

    private static string CollectAllAgentText()
    {
        var sb = new StringBuilder();
        foreach (var (_, text) in EnumerateAllAgentTexts())
            sb.AppendLine(text);
        return sb.ToString();
    }

    private static void AssertGuidanceText(string id, string text, bool requireRoutingSemantics)
    {
        foreach (var pattern in ForbiddenPatterns)
            Require(
                !text.Contains(pattern, StringComparison.OrdinalIgnoreCase),
                "forbidden-pattern", id, $"pattern={Truncate(pattern)}");

        Require(!OfficeCliCommandRegex.IsMatch(text), "officecli-command", id);

        var prose = JsonFenceRegex.Replace(text, string.Empty);
        var option = CliOptionTokenRegex.Match(prose);
        Require(!option.Success, "option-in-prose", id, $"option={Truncate(option.Value)}");
        var shellEscape = ShellEscapeRegex.Match(prose);
        Require(!shellEscape.Success, "shell-escape", id, $"value={Truncate(shellEscape.Value)}");

        foreach (Match m in UnknownOfficeTokenRegex.Matches(text))
            Require(RoleTokens.Contains(m.Value), "unknown-tool-token", id, $"token={m.Value}");

        // Agent 说明面禁止任何 officecli / OfficeCLI 字样（含命令方言与产品名），避免诱导向 CLI
        Require(
            !text.Contains("officecli", StringComparison.OrdinalIgnoreCase),
            "forbidden-product-name", id);

        foreach (var fence in ExtractJsonFences(text))
            AssertExactToolEnvelope(id, fence);

        if (requireRoutingSemantics)
        {
            Require(text.Contains("不是事务", StringComparison.Ordinal), "missing-nontransaction", id);
            Require(text.Contains("8192", StringComparison.Ordinal), "missing-overflow-limit", id);
            Require(text.Contains("outputFile", StringComparison.Ordinal), "missing-output-file", id);
        }
    }

    private static void AssertExactToolEnvelope(string id, string json)
    {
        JsonDocument doc;
        try
        {
            doc = JsonDocument.Parse(json);
        }
        catch (JsonException)
        {
            throw RuleError("invalid-json", id);
        }

        using (doc)
        {
            var root = doc.RootElement;
            Require(root.ValueKind == JsonValueKind.Object, "envelope-root", id);
            var rootKeys = root.EnumerateObject()
                .Select(property => property.Name)
                .OrderBy(name => name, StringComparer.Ordinal);
            Require(
                rootKeys.SequenceEqual(
                    new[] { "arguments", "tool" }.OrderBy(name => name, StringComparer.Ordinal)),
                "envelope-exact-shape", id);
            Require(
                root.TryGetProperty("tool", out var toolElement)
                && toolElement.ValueKind == JsonValueKind.String
                && toolElement.GetString() is { Length: > 0 },
                "missing-tool", id);
            var tool = toolElement.GetString()!;
            Require(RoleTokens.Contains(tool), "unknown-tool", id, $"tool={tool}");
            Require(
                root.TryGetProperty("arguments", out var args)
                && args.ValueKind == JsonValueKind.Object,
                "arguments-object", id);
            var keys = args.EnumerateObject()
                .Select(p => p.Name)
                .OrderBy(n => n, StringComparer.Ordinal)
                .ToArray();

            if (tool == GuidanceSurface.BatchTool)
            {
                Require(!args.TryGetProperty("commands", out _), "commands-string", id);
                var expected = new[] { "file", "operations", "stop_on_error" }
                    .OrderBy(n => n, StringComparer.Ordinal);
                Require(keys.SequenceEqual(expected), "batch-exact-shape", id);
                Require(
                    args.TryGetProperty("file", out var file)
                    && file.ValueKind == JsonValueKind.String,
                    "batch-file", id);
                Require(
                    args.TryGetProperty("stop_on_error", out var stop)
                    && stop.ValueKind is JsonValueKind.True or JsonValueKind.False,
                    "batch-stop-on-error", id);
                Require(
                    args.TryGetProperty("operations", out var ops)
                    && ops.ValueKind == JsonValueKind.Array,
                    "batch-operations-array", id);
                Require(ops.GetArrayLength() > 0, "empty-operations", id);
                foreach (var op in ops.EnumerateArray())
                    AssertBatchOperation(id, op);
            }
            else if (tool == GuidanceSurface.RunTool)
            {
                var expected = new[] { "command_arguments", "command_name" }
                    .OrderBy(n => n, StringComparer.Ordinal);
                Require(keys.SequenceEqual(expected), "run-exact-shape", id);
                Require(
                    args.TryGetProperty("command_name", out var commandName)
                    && commandName.ValueKind == JsonValueKind.String
                    && commandName.GetString() is { Length: > 0 },
                    "run-command-name", id);
                Require(
                    !string.Equals(commandName.GetString(), "batch", StringComparison.OrdinalIgnoreCase),
                    "run-batch", id);
                Require(
                    args.TryGetProperty("command_arguments", out var commandArguments)
                    && commandArguments.ValueKind == JsonValueKind.Array,
                    "run-command-arguments", id);
                foreach (var argument in commandArguments.EnumerateArray())
                    Require(argument.ValueKind == JsonValueKind.String, "run-argument-string", id);
            }
            else
            {
                // O-R02：Help/LoadSkill 仅允许 command_arguments[]，拒绝 format/topic/name 旧 DTO
                Require(
                    keys.SequenceEqual(new[] { "command_arguments" }, StringComparer.Ordinal),
                    "help-load-exact-shape", id, $"tool={tool}");
                Require(
                    args.TryGetProperty("command_arguments", out var commandArguments)
                    && commandArguments.ValueKind == JsonValueKind.Array,
                    "help-load-command-arguments", id);
                foreach (var argument in commandArguments.EnumerateArray())
                    Require(argument.ValueKind == JsonValueKind.String, "help-load-argument-string", id);
            }
        }
    }

    [Fact]
    public void AgentGuidance_HelpLoadSkillLegacyDtoShapesAreRejectedByExactShapeGate()
    {
        // 负向 fixture：旧 {format,topic}/{name} 必须被 exact-shape 门禁拒绝
        foreach (var legacy in new[]
                 {
                     """{"tool":"{{OFFICE_HELP_TOOL}}","arguments":{"format":"pptx","topic":"slide"}}""",
                     """{"tool":"{{OFFICE_LOAD_SKILL_TOOL}}","arguments":{"name":"pptx"}}""",
                 })
        {
            AssertRule(
                "help-load-exact-shape", "legacy", () => AssertExactToolEnvelope("legacy", legacy));
        }

        // 正例：仅 command_arguments
        AssertExactToolEnvelope("positive-help",
            """{"tool":"{{OFFICE_HELP_TOOL}}","arguments":{"command_arguments":["pptx","slide"]}}""");
        AssertExactToolEnvelope("positive-load",
            """{"tool":"{{OFFICE_LOAD_SKILL_TOOL}}","arguments":{"command_arguments":["pptx"]}}""");
    }

    private static void AssertBatchOperation(string id, JsonElement op)
    {
        Require(op.ValueKind == JsonValueKind.Object, "operation-object", id);
        foreach (var prop in op.EnumerateObject())
        {
            if (prop.NameEquals("op"))
                throw new InvalidOperationException($"rule:op-alias id={id}");
            if (!AllowedOperationFields.Contains(prop.Name))
                throw new InvalidOperationException($"rule:unknown-field id={id} field={prop.Name}");
        }

        Require(
            op.TryGetProperty("command", out var command)
            && command.ValueKind == JsonValueKind.String
            && command.GetString() is { Length: > 0 },
            "missing-command", id);
        var cmd = command.GetString()!;
        if (ReadHeavyCommands.Contains(cmd))
            throw new InvalidOperationException($"rule:read-heavy-batch id={id} cmd={cmd}");
        if (!BatchMutationCommands.Contains(cmd))
            throw new InvalidOperationException($"rule:non-mutation-batch id={id} cmd={cmd}");

        if (op.TryGetProperty("props", out var props))
        {
            if (props.ValueKind == JsonValueKind.Null)
                throw new InvalidOperationException($"rule:props-null id={id}");
            if (props.ValueKind != JsonValueKind.Object)
                throw new InvalidOperationException($"rule:props-not-object id={id}");
            foreach (var p in props.EnumerateObject())
            {
                if (p.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Object or JsonValueKind.Array)
                    throw new InvalidOperationException($"rule:props-compound id={id} key={p.Name}");
            }
        }

        if (SchemaHelpAgentRenderer.HasAmbiguousPositionFields(op))
            throw new InvalidOperationException($"rule:ambiguous-position id={id}");
        if (SchemaHelpAgentRenderer.HasAddFromAndProps(op))
            throw new InvalidOperationException($"rule:add-from-props id={id}");
    }

    private static void AssertBatchOperationsIndependent(string id, JsonElement operations)
    {
        var createdPrefixes = new List<string>();
        var createdNames = new List<string>();
        var writeFootprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var createdAbstractNumbering = false;
        var createdNumberingInstance = false;
        var index = 0;
        foreach (var operation in operations.EnumerateArray())
        {
            foreach (var field in new[] { "path", "parent" })
            {
                if (!operation.TryGetProperty(field, out var target)
                    || target.ValueKind != JsonValueKind.String)
                    continue;
                var value = target.GetString()!;
                if (createdPrefixes.Any(prefix => IsSameOrDescendant(value, prefix)))
                    throw new InvalidOperationException(
                        $"rule:batch-dependency id={id} operation={index} target={value}");
            }

            if (createdNames.Any(name => OperationConsumesName(operation, name)))
                throw new InvalidOperationException(
                    $"rule:batch-name-dependency id={id} operation={index}");

            if (operation.TryGetProperty("props", out var props)
                && props.ValueKind == JsonValueKind.Object
                && props.TryGetProperty("protection", out _)
                && operations.GetArrayLength() > 1)
                throw new InvalidOperationException(
                    $"rule:batch-protection-must-be-final-run id={id} operation={index}");

            if (operation.TryGetProperty("props", out props)
                && props.ValueKind == JsonValueKind.Object)
            {
                if (createdAbstractNumbering && props.TryGetProperty("abstractNumId", out _))
                    throw new InvalidOperationException(
                        $"rule:batch-generated-id-dependency id={id} operation={index}");
                if (createdNumberingInstance && props.TryGetProperty("numId", out _))
                    throw new InvalidOperationException(
                        $"rule:batch-generated-id-dependency id={id} operation={index}");

                if (operation.TryGetProperty("path", out var writePath)
                    && writePath.ValueKind == JsonValueKind.String)
                {
                    foreach (var prop in props.EnumerateObject())
                    {
                        var footprint = $"{writePath.GetString()}\0{prop.Name}";
                        if (!writeFootprints.Add(footprint))
                            throw new InvalidOperationException(
                                $"rule:batch-overlapping-write id={id} operation={index}");
                    }
                }
            }

            if (operation.TryGetProperty("command", out var command)
                && command.GetString()?.Equals("add", StringComparison.OrdinalIgnoreCase) == true)
            {
                if (TryInferCreatedPrefix(operation) is { } prefix)
                    createdPrefixes.Add(prefix);
                if (operation.TryGetProperty("props", out var addProps)
                    && addProps.ValueKind == JsonValueKind.Object
                    && addProps.TryGetProperty("name", out var name)
                    && name.ValueKind == JsonValueKind.String
                    && name.GetString() is { Length: > 0 } nameValue)
                    createdNames.Add(nameValue);

                if (operation.TryGetProperty("type", out var createdType)
                    && createdType.ValueKind == JsonValueKind.String)
                {
                    var typeName = createdType.GetString()!;
                    createdAbstractNumbering |= typeName.Equals(
                        "abstractnum", StringComparison.OrdinalIgnoreCase);
                    createdNumberingInstance |= typeName.Equals(
                        "num", StringComparison.OrdinalIgnoreCase)
                        || typeName.Equals("numberinginstance", StringComparison.OrdinalIgnoreCase);
                }
            }

            index++;
        }
    }

    private static string? TryInferCreatedPrefix(JsonElement operation)
    {
        if (!operation.TryGetProperty("parent", out var parent)
            || parent.ValueKind != JsonValueKind.String
            || !operation.TryGetProperty("type", out var type)
            || type.ValueKind != JsonValueKind.String)
            return null;

        var segment = type.GetString()!.ToLowerInvariant() switch
        {
            "paragraph" or "para" => "p",
            "run" => "r",
            "table" => "tbl",
            "row" => "tr",
            "cell" => "tc",
            "abstractnum" => "abstractNum",
            "numberinginstance" or "num" => "num",
            "namedrange" => "namedrange",
            var value => value,
        };
        var parentPath = parent.GetString()!.TrimEnd('/');
        return parentPath.Length == 0 ? $"/{segment}" : $"{parentPath}/{segment}";
    }

    private static bool IsSameOrDescendant(string path, string prefix)
    {
        if (!path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return false;
        if (path.Length == prefix.Length) return true;
        return path[prefix.Length] is '/' or '[';
    }

    private static bool OperationConsumesName(JsonElement operation, string name)
    {
        foreach (var field in new[] { "path", "parent", "from", "to", "path2" })
        {
            if (operation.TryGetProperty(field, out var value)
                && value.ValueKind == JsonValueKind.String
                && value.GetString()!.Contains(name, StringComparison.Ordinal))
                return true;
        }

        if (!operation.TryGetProperty("props", out var props)
            || props.ValueKind != JsonValueKind.Object)
            return false;
        foreach (var prop in props.EnumerateObject())
        {
            if (prop.NameEquals("name")) continue;
            if (prop.Value.ValueKind == JsonValueKind.String
                && prop.Value.GetString()!.Contains(name, StringComparison.Ordinal))
                return true;
        }
        return false;
    }

    private static List<string> ExtractJsonFences(string markdown)
    {
        var list = new List<string>();
        foreach (Match match in JsonFenceRegex.Matches(markdown))
            list.Add(match.Groups[1].Value);
        return list;
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "officecli.slnx")))
            directory = directory.Parent;
        return directory?.FullName
            ?? throw new InvalidOperationException("Could not locate officecli.slnx");
    }

    private static string Truncate(string value) =>
        value.Length <= 64 ? value : value[..64] + "…";

    private static void Require(bool condition, string rule, string id, string? detail = null)
    {
        if (!condition)
            throw RuleError(rule, id, detail);
    }

    private static InvalidOperationException RuleError(
        string rule, string id, string? detail = null) =>
        new($"rule:{rule} id={id}{(string.IsNullOrEmpty(detail) ? "" : " " + detail)}");

    private static void AssertRule(string rule, Action action)
        => AssertRule(rule, "neg", action);

    private static void AssertRule(string rule, string id, Action action)
    {
        var exception = Assert.Throws<InvalidOperationException>(action);
        Assert.Contains($"rule:{rule}", exception.Message, StringComparison.Ordinal);
        Assert.Contains($"id={id}", exception.Message, StringComparison.Ordinal);
    }
}
