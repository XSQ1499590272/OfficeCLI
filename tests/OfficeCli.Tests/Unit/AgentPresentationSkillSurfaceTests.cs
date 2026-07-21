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
/// Phase-0 presentation Agent skills (O03): morph-ppt, morph-ppt-3d, pitch-deck.
/// Author: xiesq, 2026-07-15
/// </summary>
[Trait("Speed", "Unit")]
public sealed class AgentPresentationSkillSurfaceTests
{
    private static readonly string[] PresentationSkills =
        ["morph-ppt", "morph-ppt-3d", "pitch-deck"];

    private static readonly string[] MorphReferences =
    [
        "references/decision-rules.md",
        "references/design-rules.md",
        "references/style-catalog.md",
    ];

    private static readonly Regex JsonFenceRegex =
        new(@"```json\s*(\{[\s\S]*?\})\s*```", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex OfficeCliCommandRegex =
        new(@"\bofficecli\s+\S+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex MorphStyleIdRegex =
        new(@"\b(?:dark|light|warm|vivid|bw|mixed)--[a-z0-9-]+\b",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

    [Fact]
    public void AgentPresentationCatalog_ExposesThreeCompleteSkillsInOrder()
    {
        var agent = SkillCatalog.ListAgentSkillNames();
        foreach (var name in PresentationSkills)
            Assert.Contains(name, agent);

        // 相对顺序与 SkillMap 一致
        var idxs = PresentationSkills.Select(n => agent.ToList().IndexOf(n)).ToList();
        Assert.Equal(idxs.OrderBy(i => i).ToList(), idxs);
        Assert.Equal(1, agent.Count(n => n == "morph-ppt"));
    }

    [Fact]
    public void AgentMorphStyleCatalog_PreservesCanonicalSourceStyleIds()
    {
        var sourceIndex = File.ReadAllText(Path.Combine(
            FindRepoRoot(), "skills", "morph-ppt", "reference", "styles", "INDEX.md"));
        var agentCatalog = SkillCatalog.LoadSkillFile(
            "morph-ppt", "references/style-catalog.md", GuidanceSurfaceKind.AgentJson);
        var sourceIds = MorphStyleIdRegex.Matches(sourceIndex)
            .Select(match => match.Value)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        Assert.NotEmpty(sourceIds);
        foreach (var id in sourceIds)
            Assert.Contains(id, agentCatalog, StringComparison.Ordinal);
    }

    [Fact]
    public void AgentPresentationSkills_ContainRequiredDomainInvariants()
    {
        var morph = SkillCatalog.LoadSkillContent("morph-ppt", GuidanceSurfaceKind.AgentJson);
        Assert.Contains("identical", morph, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("transition", morph, StringComparison.OrdinalIgnoreCase);

        var morph3d = SkillCatalog.LoadSkillContent("morph-ppt-3d", GuidanceSurfaceKind.AgentJson);
        Assert.Contains("glb", morph3d, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("approval", morph3d, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("degrade", morph3d, StringComparison.OrdinalIgnoreCase);

        var pitch = SkillCatalog.LoadSkillContent("pitch-deck", GuidanceSurfaceKind.AgentJson);
        Assert.Contains("Do not invent", pitch, StringComparison.Ordinal);
        Assert.Contains("traction", pitch, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AgentPresentationSkills_AllJsonFencesParse()
    {
        foreach (var name in PresentationSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            AssertJsonFences(text, name);
        }

        foreach (var path in MorphReferences)
            AssertJsonFences(
                SkillCatalog.LoadSkillFile("morph-ppt", path, GuidanceSurfaceKind.AgentJson),
                path);
    }

    [Fact]
    public void AgentMorphReferences_ExactlyMatchManifest()
    {
        var listed = SkillCatalog.ListSkillFiles("morph-ppt", GuidanceSurfaceKind.AgentJson);
        Assert.Equal(MorphReferences, listed);
        foreach (var path in MorphReferences)
        {
            var content = SkillCatalog.LoadSkillFile("morph-ppt", path, GuidanceSurfaceKind.AgentJson);
            Assert.False(string.IsNullOrWhiteSpace(content));
        }
    }

    [Fact]
    public void AgentMorphPath_DeniesDefaultScriptsAssetsAndUnlistedStyles()
    {
        var denied = new[]
        {
            "reference/styles/dark--premium-navy/style.md",
            "reference/decision-rules.md",
            "build.sh",
            "model.glb",
            "preview.png",
            "references/../SKILL.md",
        };
        foreach (var path in denied)
        {
            var (code, stdout, _) = Dispatch("morph-ppt", path);
            Assert.NotEqual(0, code);
            Assert.DoesNotContain("Scene Actors", stdout, StringComparison.Ordinal);
            Assert.DoesNotContain("build.sh", stdout, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AgentPresentationSkills_ExcludeCliShellAndHostNames()
    {
        var sb = new StringBuilder();
        foreach (var name in PresentationSkills)
            sb.Append(SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson));
        foreach (var path in MorphReferences)
            sb.Append(SkillCatalog.LoadSkillFile("morph-ppt", path, GuidanceSurfaceKind.AgentJson));
        AssertNoForbidden(sb.ToString());
    }

    [Fact]
    public void AgentPresentationBatchExamples_StayWithinMutationFastPath()
    {
        foreach (var name in PresentationSkills)
            AssertBatchAndRunRouting(SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson), name);
    }

    [Fact]
    public void AgentPresentationRunExamples_DoNotInvokeBatch()
    {
        foreach (var name in PresentationSkills)
        {
            foreach (var fence in ExtractJsonFences(
                         SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson)))
            {
                using var doc = JsonDocument.Parse(fence);
                if (doc.RootElement.GetProperty("tool").GetString() != GuidanceSurface.RunTool)
                    continue;
                Assert.NotEqual(
                    "batch",
                    doc.RootElement.GetProperty("arguments").GetProperty("command_name").GetString());
            }
        }
    }

    [Fact]
    public void AgentPresentationBatchExamples_RejectAmbiguousPositionsAndFromProps()
    {
        using var badPos = JsonDocument.Parse(
            """{"command":"move","path":"/a","index":1,"after":"/b"}""");
        Assert.True(SchemaHelpAgentRenderer.HasAmbiguousPositionFields(badPos.RootElement));

        using var badFrom = JsonDocument.Parse(
            """{"command":"add","parent":"/s","type":"picture","from":"/x.png","props":{"w":"1"}}""");
        Assert.True(SchemaHelpAgentRenderer.HasAddFromAndProps(badFrom.RootElement));

        foreach (var name in PresentationSkills)
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
    public void AgentPresentationSkills_RequirePartialSuccessBoundaryAndIndependentVisualValidation()
    {
        foreach (var name in PresentationSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            Assert.Contains("不是事务", text, StringComparison.Ordinal);
            Assert.Contains("不回滚", text, StringComparison.Ordinal);
            Assert.Contains("discardable", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("validate", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("不会自动创建草稿", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AgentPresentationSkills_ExplainUnnormalizedBatchOverflow()
    {
        foreach (var name in PresentationSkills)
        {
            var text = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.AgentJson);
            Assert.Contains("8192", text, StringComparison.Ordinal);
            Assert.Contains("outputFile", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void DefaultPresentationSkillsRemainUnchanged()
    {
        foreach (var name in PresentationSkills)
        {
            var cli = SkillCatalog.LoadSkillContent(name, GuidanceSurfaceKind.Cli);
            Assert.DoesNotContain(GuidanceSurface.BatchTool, cli, StringComparison.Ordinal);
            Assert.Contains("officecli", cli, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void AgentMorphCorpusStaysWithinResourceBudget()
    {
        long total = 0;
        void Check(string label, string content)
        {
            var bytes = Encoding.UTF8.GetByteCount(content);
            Assert.True(bytes <= 128 * 1024, $"{label} too large: {bytes}");
            total += bytes;
        }

        Check("morph-ppt", SkillCatalog.LoadSkillContent("morph-ppt", GuidanceSurfaceKind.AgentJson));
        Check("morph-ppt-3d", SkillCatalog.LoadSkillContent("morph-ppt-3d", GuidanceSurfaceKind.AgentJson));
        Check("pitch-deck", SkillCatalog.LoadSkillContent("pitch-deck", GuidanceSurfaceKind.AgentJson));
        foreach (var path in MorphReferences)
            Check(path, SkillCatalog.LoadSkillFile("morph-ppt", path, GuidanceSurfaceKind.AgentJson));

        Assert.True(total <= 384 * 1024, $"morph corpus too large: {total}");
    }

    private static void AssertBatchAndRunRouting(string text, string label)
    {
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
                Assert.True(ops.GetArrayLength() >= 3, $"{label} Batch < 3");
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

        Assert.True(sawBatch, $"{label} missing Batch example");
    }

    private static void AssertJsonFences(string text, string label)
    {
        foreach (var fence in ExtractJsonFences(text))
        {
            using var doc = JsonDocument.Parse(fence);
            Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);
            var tool = doc.RootElement.GetProperty("tool").GetString();
            Assert.Contains(tool, new[]
            {
                GuidanceSurface.HelpTool,
                GuidanceSurface.LoadSkillTool,
                GuidanceSurface.BatchTool,
                GuidanceSurface.RunTool,
            });
            Assert.True(doc.RootElement.TryGetProperty("arguments", out _), label);
        }
    }

    private static void AssertNoForbidden(string text)
    {
        Assert.DoesNotContain("tool.file.office.", text, StringComparison.Ordinal);
        Assert.DoesNotContain("--commands", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("```bash", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("```python", text, StringComparison.OrdinalIgnoreCase);
        Assert.False(OfficeCliCommandRegex.IsMatch(text));
    }

    private static List<string> ExtractJsonFences(string markdown)
    {
        var list = new List<string>();
        foreach (Match m in JsonFenceRegex.Matches(markdown))
            list.Add(m.Groups[1].Value);
        return list;
    }

    private static (int Code, string Stdout, string Stderr) Dispatch(string skill, string path)
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var code = LoadSkillCli.Dispatch(
            ["--surface", "agent-json", skill, "--path", path], stdout, stderr);
        return (code, stdout.ToString(), stderr.ToString());
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "officecli.slnx")))
            directory = directory.Parent;
        return directory?.FullName
            ?? throw new InvalidOperationException("Could not locate officecli.slnx");
    }
}
