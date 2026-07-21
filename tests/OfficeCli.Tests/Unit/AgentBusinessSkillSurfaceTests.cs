// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Author: xiesq
// Created: 2026-07-15

using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeCli.Core;
using OfficeCli.Help;

namespace OfficeCli.Tests.Unit;

/// <summary>
/// Phase-0 business Agent skills (O04): academic-paper, data-dashboard, financial-model.
/// Author: xiesq, 2026-07-15
/// </summary>
[Trait("Speed", "Unit")]
public sealed class AgentBusinessSkillSurfaceTests
{
    private static readonly string[] BusinessSkills =
        ["academic-paper", "data-dashboard", "financial-model"];

    private static readonly Regex JsonFenceRegex =
        new(@"```json\s*(\{[\s\S]*?\})\s*```", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex OfficeCliCommandRegex =
        new(@"\bofficecli\s+\S+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    [Fact]
    public void AgentBusinessCatalog_ExposesThreeCompleteSkillsInOrder()
    {
        var agent = SkillCatalog.ListAgentSkillNames();
        foreach (var name in BusinessSkills)
            Assert.Contains(name, agent);

        var idxs = BusinessSkills.Select(n => agent.ToList().IndexOf(n)).ToList();
        Assert.Equal(idxs.OrderBy(i => i).ToList(), idxs);
    }

    [Fact]
    public void AgentBusinessSkills_ContainRequiredTruthAndValidationRules()
    {
        var paper = SkillCatalog.LoadSkillContent("academic-paper", GuidanceSurfaceKind.AgentJson);
        Assert.Contains("Do not invent", paper, StringComparison.Ordinal);
        Assert.Contains("citation", paper, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("validate", paper, StringComparison.OrdinalIgnoreCase);

        var dash = SkillCatalog.LoadSkillContent("data-dashboard", GuidanceSurfaceKind.AgentJson);
        Assert.Contains("metric", dash, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Do not invent", dash, StringComparison.Ordinal);
        Assert.Contains("validate", dash, StringComparison.OrdinalIgnoreCase);

        var model = SkillCatalog.LoadSkillContent("financial-model", GuidanceSurfaceKind.AgentJson);
        Assert.Contains("assumption", model, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("formula", model, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#REF!", model, StringComparison.Ordinal);
        Assert.Contains("validate", model, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AgentBusinessSkills_AllJsonFencesParse()
    {
        foreach (var name in BusinessSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            foreach (var fence in ExtractJsonFences(text))
            {
                using var doc = JsonDocument.Parse(fence);
                Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);
                var tool = doc.RootElement.GetProperty("tool").GetString();
                Assert.Contains(tool, new[]
                {
                    GuidanceSurface.BatchTool,
                    GuidanceSurface.RunTool,
                    GuidanceSurface.HelpTool,
                    GuidanceSurface.LoadSkillTool,
                });
                Assert.True(doc.RootElement.TryGetProperty("arguments", out _));
            }
        }
    }

    [Fact]
    public void AgentBusinessSkills_BatchExamplesUseArraysAndScalarProps()
    {
        foreach (var name in BusinessSkills)
        {
            foreach (var fence in ExtractJsonFences(
                         SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson)))
            {
                using var doc = JsonDocument.Parse(fence);
                if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.BatchTool)
                    continue;
                var args = doc.RootElement.GetProperty("arguments");
                Assert.Equal(JsonValueKind.Array, args.GetProperty("operations").ValueKind);
                Assert.False(args.TryGetProperty("commands", out _));
                foreach (var op in args.GetProperty("operations").EnumerateArray())
                {
                    if (!op.TryGetProperty("props", out var props)) continue;
                    Assert.Equal(JsonValueKind.Object, props.ValueKind);
                    foreach (var p in props.EnumerateObject())
                    {
                        Assert.True(
                            p.Value.ValueKind is JsonValueKind.String
                                or JsonValueKind.Number
                                or JsonValueKind.True
                                or JsonValueKind.False,
                            $"{name} props.{p.Name} must be scalar");
                    }
                }
            }
        }
    }

    [Fact]
    public void AgentBusinessSkills_RouteBatchOnlyForIndependentMutations()
    {
        foreach (var name in BusinessSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            var sawBatch = false;
            foreach (var fence in ExtractJsonFences(text))
            {
                using var doc = JsonDocument.Parse(fence);
                var tool = doc.RootElement.GetProperty("tool").GetString();
                var args = doc.RootElement.GetProperty("arguments");
                if (tool == GuidanceSurface.BatchTool)
                {
                    sawBatch = true;
                    var ops = args.GetProperty("operations");
                    Assert.True(ops.GetArrayLength() >= 3);
                    foreach (var op in ops.EnumerateArray())
                    {
                        var cmd = op.GetProperty("command").GetString()!;
                        Assert.Contains(cmd, SchemaHelpAgentRenderer.BatchMutationCommands);
                        Assert.DoesNotContain(cmd, SchemaHelpAgentRenderer.ReadOrCheckCommands);
                    }
                }
                else if (tool == GuidanceSurface.RunTool)
                {
                    Assert.NotEqual("batch", args.GetProperty("command_name").GetString());
                }
            }

            Assert.True(sawBatch, $"{name} missing Batch");
        }
    }

    [Fact]
    public void AgentBusinessSkills_RunRejectsOnlyBatchCommand()
    {
        var cmds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in BusinessSkills)
        {
            foreach (var fence in ExtractJsonFences(
                         SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson)))
            {
                using var doc = JsonDocument.Parse(fence);
                if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.RunTool)
                    continue;
                var cmd = doc.RootElement.GetProperty("arguments").GetProperty("command_name").GetString()!;
                Assert.NotEqual("batch", cmd);
                cmds.Add(cmd);
            }
        }

        Assert.Contains("validate", cmds);
        Assert.Contains("get", cmds);
    }

    [Fact]
    public void AgentBusinessSkills_EnforceBatchCrossFieldRules()
    {
        using var bad = JsonDocument.Parse(
            """{"command":"add","parent":"/a","type":"p","index":1,"before":"/b"}""");
        Assert.True(SchemaHelpAgentRenderer.HasAmbiguousPositionFields(bad.RootElement));

        foreach (var name in BusinessSkills)
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
    public void AgentBusinessSkillPath_DeniesAllReferencesWithoutFallback()
    {
        var probes = new[]
        {
            "reference/foo.md",
            "../officecli-docx/SKILL.md",
            "/tmp/x.md",
            "a\\b.md",
            "script.py",
        };
        foreach (var name in BusinessSkills)
        {
            foreach (var path in probes)
            {
                using var stdout = new StringWriter();
                using var stderr = new StringWriter();
                var code = LoadSkillCli.Dispatch(
                    ["--surface", "agent-json", name, "--path", path], stdout, stderr);
                Assert.NotEqual(0, code);
                Assert.DoesNotContain("scene layer", stdout.ToString(), StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    [Fact]
    public void AgentBusinessSkills_ExcludeCliShellAndHostNames()
    {
        var sb = new StringBuilder();
        foreach (var name in BusinessSkills)
            sb.Append(SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson));
        var text = sb.ToString();
        Assert.DoesNotContain("tool.file.office.", text, StringComparison.Ordinal);
        Assert.DoesNotContain("--commands", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("```bash", text, StringComparison.OrdinalIgnoreCase);
        Assert.False(OfficeCliCommandRegex.IsMatch(text));
    }

    [Fact]
    public void AgentBusinessExamples_ReadAfterWrite()
    {
        foreach (var name in BusinessSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            Assert.Contains("discardable", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("validate", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("save", text, StringComparison.OrdinalIgnoreCase);
            var tools = ExtractJsonFences(text)
                .Select(f =>
                {
                    using var d = JsonDocument.Parse(f);
                    return d.RootElement.GetProperty("tool").GetString();
                })
                .ToHashSet();
            Assert.Contains(GuidanceSurface.BatchTool, tools);
            Assert.Contains(GuidanceSurface.RunTool, tools);
        }
    }

    [Fact]
    public void AgentBusinessSkills_ExplainBatchFailureAndOverflowBoundaries()
    {
        foreach (var name in BusinessSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            Assert.Contains("不是事务", text, StringComparison.Ordinal);
            Assert.Contains("不回滚", text, StringComparison.Ordinal);
            Assert.Contains("8192", text, StringComparison.Ordinal);
            Assert.Contains("outputFile", text, StringComparison.Ordinal);
            Assert.Contains("不会自动创建草稿", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void DefaultBusinessSkillsRemainUnchanged()
    {
        foreach (var name in BusinessSkills)
        {
            var cli = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.Cli);
            Assert.DoesNotContain(GuidanceSurface.BatchTool, cli, StringComparison.Ordinal);
            Assert.Contains("officecli", cli, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void AgentBusinessCorpusStaysWithinResourceBudget()
    {
        foreach (var name in BusinessSkills)
        {
            var bytes = Encoding.UTF8.GetByteCount(
                SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson));
            Assert.True(bytes <= 128 * 1024, $"{name} too large: {bytes}");
        }
    }

    private static List<string> ExtractJsonFences(string markdown)
    {
        var list = new List<string>();
        foreach (Match m in JsonFenceRegex.Matches(markdown))
            list.Add(m.Groups[1].Value);
        return list;
    }
}
