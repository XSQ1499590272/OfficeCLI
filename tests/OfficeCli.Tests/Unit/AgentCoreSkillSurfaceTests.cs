// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Author: xiesq
// Created: 2026-07-15

using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeCli.Core;
using OfficeCli.Help;

namespace OfficeCli.Tests.Unit;

/// <summary>
/// Phase-0 Agent Skill loader and four core skills (O02).
/// Author: xiesq, 2026-07-15
/// </summary>
[Trait("Speed", "Unit")]
public sealed class AgentCoreSkillSurfaceTests
{
    // 与 SkillMap 插入顺序中四个核心 Skill 的相对顺序一致
    private static readonly string[] CoreSkills = ["pptx", "word", "excel", "word-form"];

    private static readonly string[] RolePlaceholders =
    [
        GuidanceSurface.HelpTool,
        GuidanceSurface.LoadSkillTool,
        GuidanceSurface.BatchTool,
        GuidanceSurface.RunTool,
    ];

    private static readonly Regex JsonFenceRegex =
        new(@"```json\s*(\{[\s\S]*?\})\s*```", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex OfficeCliCommandRegex =
        new(@"\bofficecli\s+\S+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

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
    public void AgentSkillParser_AcceptsSurfaceBeforeOrAfterPath()
    {
        // catalog（必须显式 agent-json）
        AssertDispatchOk(["--surface", "agent-json"], GuidanceSurfaceKind.AgentJson, out var catalog);
        Assert.Contains("## word", catalog, StringComparison.Ordinal);

        // skill only
        AssertDispatchOk(["word", "--surface", "agent-json"], out var skillA);
        AssertDispatchOk(["--surface", "agent-json", "word"], out var skillB);
        Assert.Equal(skillA, skillB);

        // skill + path orderings（空 references → 均应失败，但解析本身要一致到 manifest 检查）
        var orders = new[]
        {
            new[] { "word", "--surface", "agent-json", "--path", "missing.md" },
            new[] { "--surface", "agent-json", "word", "--path", "missing.md" },
            new[] { "word", "--path", "missing.md", "--surface", "agent-json" },
            new[] { "--path", "missing.md", "--surface", "agent-json", "word" },
        };
        foreach (var tokens in orders)
        {
            var (code, _, err) = Dispatch(tokens);
            Assert.NotEqual(0, code);
            Assert.Contains("not declared", err, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void AgentSkillParser_RejectsDuplicateMissingAndUnknownOptions()
    {
        Assert.NotEqual(0, Dispatch(["--surface", "agent-json", "--surface", "cli"]).Code);
        Assert.NotEqual(0, Dispatch(["--surface", "", "--surface", "agent-json"]).Code);
        Assert.NotEqual(0, Dispatch(["--surface"]).Code);
        Assert.NotEqual(0, Dispatch(["--surface=agent-json"]).Code);
        Assert.NotEqual(0, Dispatch(["word", "--path", ""]).Code);
        Assert.NotEqual(0, Dispatch(["word", "--path", "", "--path", "missing.md"]).Code);
        Assert.NotEqual(0, Dispatch(["--unknown", "word"]).Code);
        Assert.NotEqual(0, Dispatch(["word", "excel", "--surface", "agent-json"]).Code);

        // 不得回退到 CLI catalog 内容
        var bad = Dispatch(["--surface=agent-json"]);
        Assert.DoesNotContain("officecli skills", bad.Stdout, StringComparison.Ordinal);
    }

    [Fact]
    public void AgentSkillCatalog_ListsOnlyCompleteAgentResourcesInStableOrder()
    {
        var agentNames = SkillCatalog.ListAgentSkillNames();
        // 四个核心必须存在；O03/O04 合入后会扩展到完整 SkillMap
        foreach (var name in CoreSkills)
            Assert.Contains(name, agentNames);

        // 核心相对顺序与 SkillMap 一致
        var mapOrder = SkillCatalog.SkillMapNamesInOrder()
            .Where(n => CoreSkills.Contains(n, StringComparer.OrdinalIgnoreCase))
            .ToList();
        var coreInAgent = agentNames
            .Where(n => CoreSkills.Contains(n, StringComparer.OrdinalIgnoreCase))
            .ToList();
        Assert.Equal(mapOrder, coreInAgent);

        var cliCatalog = SkillCatalog.BuildSkillCatalog(GuidanceSurfaceKind.Cli);
        foreach (var name in SkillCatalog.SkillMapNamesInOrder())
            Assert.Contains($"## {name}", cliCatalog, StringComparison.Ordinal);

        // CLI catalog 可保留「# officecli skills」；Agent catalog 禁止任何 officecli 字样
        Assert.Contains("# officecli skills", cliCatalog, StringComparison.Ordinal);
        var agentCatalog = SkillCatalog.BuildSkillCatalog(GuidanceSurfaceKind.AgentJson);
        Assert.Contains("# Agent skills", agentCatalog, StringComparison.Ordinal);
        Assert.DoesNotContain("officecli", agentCatalog, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AgentCoreSkills_LoadIndependentToolJsonGuidance()
    {
        foreach (var name in CoreSkills)
        {
            var agent = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            var cli = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.Cli);
            Assert.NotEqual(cli, agent);
            Assert.Contains("{{OFFICE_", agent, StringComparison.Ordinal);
            Assert.Contains("不是事务", agent, StringComparison.Ordinal);
            Assert.Contains("Inspect", agent, StringComparison.OrdinalIgnoreCase);
            Assert.NotEmpty(ExtractJsonFences(agent));
        }
    }

    [Fact]
    public void AgentCoreSkills_UseOnlyRolePlaceholders()
    {
        foreach (var name in CoreSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            foreach (var token in RolePlaceholders)
                Assert.Contains(token, text, StringComparison.Ordinal);
            Assert.DoesNotContain("tool.file.office.", text, StringComparison.Ordinal);
            AssertNoForbiddenSyntax(text);
        }
    }

    [Fact]
    public void AgentCoreSkills_RouteBatchOnlyForIndependentMutations()
    {
        foreach (var name in CoreSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            var batchOps = new List<JsonElement>();
            foreach (var fence in ExtractJsonFences(text))
            {
                using var doc = JsonDocument.Parse(fence);
                var tool = doc.RootElement.GetProperty("tool").GetString();
                var args = doc.RootElement.GetProperty("arguments");
                if (tool == GuidanceSurface.BatchTool)
                {
                    var ops = args.GetProperty("operations");
                    Assert.True(ops.GetArrayLength() >= 3, $"{name} Batch example too short");
                    foreach (var op in ops.EnumerateArray())
                    {
                        var cmd = op.GetProperty("command").GetString()!;
                        Assert.Contains(cmd, SchemaHelpAgentRenderer.BatchMutationCommands);
                        Assert.DoesNotContain(cmd, SchemaHelpAgentRenderer.ReadOrCheckCommands);
                        batchOps.Add(op.Clone());
                    }
                }
                else if (tool == GuidanceSurface.RunTool)
                {
                    var cmd = args.GetProperty("command_name").GetString();
                    Assert.NotEqual("batch", cmd);
                }
            }

            Assert.NotEmpty(batchOps);
        }
    }

    [Fact]
    public void AgentCoreSkills_KeepRunAvailableForNonBatchCommands()
    {
        var runCommands = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in CoreSkills)
        {
            foreach (var fence in ExtractJsonFences(
                         SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson)))
            {
                using var doc = JsonDocument.Parse(fence);
                if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.RunTool)
                    continue;
                var cmd = doc.RootElement.GetProperty("arguments").GetProperty("command_name").GetString()!;
                Assert.NotEqual("batch", cmd);
                runCommands.Add(cmd);
            }
        }

        Assert.Contains("validate", runCommands);
        Assert.Contains("get", runCommands);
    }

    [Fact]
    public void AgentCoreSkills_EnforceBatchCrossFieldRules()
    {
        foreach (var name in CoreSkills)
        {
            foreach (var fence in ExtractJsonFences(
                         SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson)))
            {
                using var doc = JsonDocument.Parse(fence);
                if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.BatchTool)
                    continue;
                foreach (var op in doc.RootElement.GetProperty("arguments").GetProperty("operations").EnumerateArray())
                {
                    Assert.False(SchemaHelpAgentRenderer.HasAmbiguousPositionFields(op));
                    Assert.False(SchemaHelpAgentRenderer.HasAddFromAndProps(op));
                }
            }
        }
    }

    [Fact]
    public void AgentCoreSkills_TeachPartialSuccessAndIndependentValidation()
    {
        foreach (var name in CoreSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            Assert.Contains("不是事务", text, StringComparison.Ordinal);
            Assert.Contains("不回滚", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("partial", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("discardable", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("validate", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("不会自动创建草稿", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AgentCoreSkills_NeverBatchReadHeavyCommands()
    {
        foreach (var name in CoreSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            Assert.Contains("8192", text, StringComparison.Ordinal);
            Assert.Contains("outputFile", text, StringComparison.Ordinal);
            foreach (var fence in ExtractJsonFences(text))
            {
                using var doc = JsonDocument.Parse(fence);
                if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.BatchTool)
                    continue;
                foreach (var op in doc.RootElement.GetProperty("arguments").GetProperty("operations").EnumerateArray())
                {
                    var cmd = op.GetProperty("command").GetString()!;
                    Assert.DoesNotContain(cmd, SchemaHelpAgentRenderer.ReadOrCheckCommands);
                }
            }
        }
    }

    [Fact]
    public void AgentSkillPath_RejectsAnythingOutsideManifestWithoutFallback()
    {
        var probes = new[]
        {
            "reference/decision-rules.md",
            "../officecli-docx/SKILL.md",
            "/etc/passwd",
            "foo\\bar.md",
            "build.sh",
            "script.py",
            "SKILL.md",
            "manifest.json",
        };
        foreach (var name in CoreSkills)
        {
            foreach (var path in probes)
            {
                var (code, stdout, _) = Dispatch(
                    ["--surface", "agent-json", name, "--path", path]);
                Assert.NotEqual(0, code);
                // 不得泄漏默认 CLI Skill 正文特征（CLI docx skill 含 Shell 段落）
                Assert.DoesNotContain("Shell 与执行规范", stdout, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void AgentSkillManifest_RejectsUnknownDuplicateAndMissingEntries()
    {
        Assert.Equal(10, SkillCatalog.ListAgentSkillNames().Count);

        const string valid = """
            {"name":"word","description":"Word guidance","trigger":"Word documents","references":["references/a.md"]}
            """;
        var complete = new HashSet<string>(StringComparer.Ordinal)
            { "SKILL.md", "references/a.md" };
        var parsed = SkillCatalog.ParseAgentManifest(
            "word", "officecli-docx", valid, complete.Contains);
        Assert.Equal("word", parsed.Name);

        var fixtures = new[]
        {
            ("unknown-field",
                """{"name":"word","description":"d","trigger":"t","references":[],"extra":true}""",
                new HashSet<string>(StringComparer.Ordinal) { "SKILL.md" }),
            ("duplicate-field",
                """{"name":"word","name":"word","description":"d","trigger":"t","references":[]}""",
                new HashSet<string>(StringComparer.Ordinal) { "SKILL.md" }),
            ("duplicate-reference",
                """{"name":"word","description":"d","trigger":"t","references":["a.md","a.md"]}""",
                new HashSet<string>(StringComparer.Ordinal) { "SKILL.md", "a.md" }),
            ("missing-description",
                """{"name":"word","trigger":"t","references":[]}""",
                new HashSet<string>(StringComparer.Ordinal) { "SKILL.md" }),
            ("missing-skill",
                """{"name":"word","description":"d","trigger":"t","references":[]}""",
                new HashSet<string>(StringComparer.Ordinal)),
            ("missing-reference",
                """{"name":"word","description":"d","trigger":"t","references":["missing.md"]}""",
                new HashSet<string>(StringComparer.Ordinal) { "SKILL.md" }),
            ("reference-extension",
                """{"name":"word","description":"d","trigger":"t","references":["run.exe"]}""",
                new HashSet<string>(StringComparer.Ordinal) { "SKILL.md", "run.exe" }),
        };

        foreach (var (category, manifest, resources) in fixtures)
        {
            var ex = Assert.Throws<SkillCatalog.AgentManifestException>(() =>
                SkillCatalog.ParseAgentManifest(
                    "word", "officecli-docx", manifest, resources.Contains));
            Assert.Equal(category, ex.Code);
            Assert.DoesNotContain("agent-skills/", ex.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AgentSkillPath_RejectsBackslashOnAllowlistedReferenceBeforeNormalization()
    {
        // O-R03：allowlist 中的正斜杠路径改用反斜杠请求时，必须因反斜杠失败，而非误读成功
        var forward = SkillCatalog.LoadSkillFile(
            "morph-ppt", "references/decision-rules.md", GuidanceSurfaceKind.AgentJson);
        Assert.False(string.IsNullOrWhiteSpace(forward));

        var ex = Assert.Throws<ArgumentException>(() =>
            SkillCatalog.LoadSkillFile(
                "morph-ppt", @"references\decision-rules.md", GuidanceSurfaceKind.AgentJson));
        Assert.Contains("forward slash", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("not declared", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DefaultSkillSurfaceKeepsExistingCliResources()
    {
        var cli = SkillCatalog.BuildSkillCatalog();
        Assert.Contains("# officecli skills", cli, StringComparison.Ordinal);
        Assert.DoesNotContain(GuidanceSurface.BatchTool, cli, StringComparison.Ordinal);

        var word = SkillCatalog.LoadSkillContent("word");
        Assert.Contains("officecli help docx", word, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AgentSkillConcurrentReadsAreDeterministic()
    {
        var results = new string[32];
        Parallel.For(0, 32, i =>
        {
            var sb = new StringBuilder();
            for (var n = 0; n < 100; n++)
            {
                foreach (var name in CoreSkills)
                    sb.Append(SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson));
            }

            results[i] = sb.ToString();
        });

        Assert.All(results, r => Assert.Equal(results[0], r));
    }

    [Fact]
    public void AgentCoreSkillResourcesAreEmbeddedWithStablePrefix()
    {
        var names = Assembly.GetAssembly(typeof(SkillCatalog))!
            .GetManifestResourceNames()
            .Where(n => n.Replace('\\', '/')
                .StartsWith("agent-skills/", StringComparison.OrdinalIgnoreCase))
            .Select(n => n.Replace('\\', '/'))
            .ToList();

        Assert.Contains(names, n => n.EndsWith("officecli-docx/manifest.json", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(names, n => n.EndsWith("officecli-docx/SKILL.md", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(names, n => n.EndsWith("officecli-xlsx/SKILL.md", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(names, n => n.EndsWith("officecli-pptx/SKILL.md", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(names, n => n.EndsWith("officecli-word-form/SKILL.md", StringComparison.OrdinalIgnoreCase));

        // catalog + four skills load without disk fallback
        Assert.Equal(0, Dispatch(["--surface", "agent-json"]).Code);
        foreach (var name in CoreSkills)
            Assert.Equal(0, Dispatch(["--surface", "agent-json", name]).Code);
    }

    private static void AssertDispatchOk(string[] tokens, out string stdout) =>
        AssertDispatchOk(tokens, null, out stdout);

    private static void AssertDispatchOk(
        string[] tokens, GuidanceSurfaceKind? expectSurfaceHint, out string stdout)
    {
        var (code, outText, err) = Dispatch(tokens);
        Assert.True(code == 0, err);
        stdout = outText;
        if (expectSurfaceHint == GuidanceSurfaceKind.AgentJson)
            Assert.Contains("agent skills", outText, StringComparison.OrdinalIgnoreCase);
    }

    private static (int Code, string Stdout, string Stderr) Dispatch(params string[] tokens)
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var code = LoadSkillCli.Dispatch(tokens, stdout, stderr);
        return (code, stdout.ToString(), stderr.ToString());
    }

    private static List<string> ExtractJsonFences(string markdown)
    {
        var list = new List<string>();
        foreach (Match match in JsonFenceRegex.Matches(markdown))
            list.Add(match.Groups[1].Value);
        return list;
    }

    private static void AssertNoForbiddenSyntax(string text)
    {
        foreach (var pattern in ForbiddenPatterns)
            Assert.DoesNotContain(pattern, text, StringComparison.OrdinalIgnoreCase);
        Assert.False(OfficeCliCommandRegex.IsMatch(text), "executable officecli command found");
    }
}
