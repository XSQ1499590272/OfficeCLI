// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Author: xiesq
// Created: 2026-07-15

using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Help;

namespace OfficeCli.Tests.Unit;

/// <summary>
/// Phase-0 Agent Help surface tests (O01).
/// Author: xiesq, 2026-07-15
/// </summary>
[Collection("Word CLI output")]
[Trait("Speed", "Unit")]
public sealed class AgentHelpSurfaceTests
{
    private static readonly Regex JsonFenceRegex =
        new(@"```json\s*(\{[\s\S]*?\})\s*```", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly string[] ForbiddenPatterns =
    {
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
    };

    // 禁止可执行 officecli <token>；Agent 说明面进一步禁止任何 officecli 子串
    private static readonly Regex OfficeCliCommandRegex =
        new(@"\bofficecli\s+\S+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    [Fact]
    public void AgentHelp_DefaultSurfaceKeepsExistingCliRenderer()
    {
        // 未传 surface：no-args 仍含 SCL 命令面与 schema 参考
        var noArgs = InvokeCaptured("help");
        Assert.Equal(0, noArgs.ExitCode);
        Assert.Contains("Description:", noArgs.Stdout, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Schema 参考", noArgs.Stdout, StringComparison.Ordinal);
        Assert.DoesNotContain(GuidanceSurface.BatchTool, noArgs.Stdout, StringComparison.Ordinal);

        // 显式 cli 与默认等价于人类 Help
        var explicitCli = InvokeCaptured("help", "--surface", "cli");
        Assert.Equal(0, explicitCli.ExitCode);
        Assert.Contains("Schema 参考", explicitCli.Stdout, StringComparison.Ordinal);

        // 单元素人类渲染仍含用法/属性
        var element = InvokeCaptured("help", "docx", "paragraph");
        Assert.Equal(0, element.ExitCode);
        Assert.Contains("paragraph", element.Stdout, StringComparison.Ordinal);
        Assert.Contains("用法：", element.Stdout, StringComparison.Ordinal);

        // --json 仍返回可解析 schema JSON（裸文档）
        var json = InvokeCaptured("help", "docx", "paragraph", "--json");
        Assert.Equal(0, json.ExitCode);
        using var doc = JsonDocument.Parse(json.Stdout);
        Assert.Equal("paragraph", doc.RootElement.GetProperty("element").GetString());
    }

    [Fact]
    public void AgentHelp_SingleAddRendersRunEnvelope()
    {
        var result = InvokeCaptured("help", "docx", "add", "paragraph", "--surface", "agent-json");
        Assert.Equal(0, result.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(result.Stderr), result.Stderr);

        var fences = ExtractJsonFences(result.Stdout);
        Assert.NotEmpty(fences);

        var sawAddRun = false;
        foreach (var fence in fences)
        {
            using var json = JsonDocument.Parse(fence);
            var root = json.RootElement;
            Assert.Equal(GuidanceSurface.RunTool, root.GetProperty("tool").GetString());
            var args = root.GetProperty("arguments");
            Assert.Equal("add", args.GetProperty("command_name").GetString());
            Assert.Equal(JsonValueKind.Array, args.GetProperty("command_arguments").ValueKind);
            Assert.DoesNotContain(
                args.EnumerateObject().Select(p => p.Name),
                name => name is "operations" or "file" or "stop_on_error");
            sawAddRun = true;
        }

        Assert.True(sawAddRun, "expected at least one Run envelope for add");
        // 单命令 Help 不得输出长度为 1 的 Batch
        Assert.DoesNotContain(GuidanceSurface.BatchTool, result.Stdout, StringComparison.Ordinal);
        AssertNoForbiddenSyntax(result.Stdout);
    }

    [Fact]
    public void AgentHelp_VerbFilterRoutesMutationAndReadsToRun()
    {
        foreach (var verb in new[] { "add", "set", "get", "query" })
        {
            var result = InvokeCaptured(
                "help", "docx", verb, "paragraph", "--surface", "agent-json");
            Assert.Equal(0, result.ExitCode);
            Assert.DoesNotContain(GuidanceSurface.BatchTool, result.Stdout, StringComparison.Ordinal);

            var fences = ExtractJsonFences(result.Stdout);
            Assert.NotEmpty(fences);
            foreach (var fence in fences)
            {
                using var json = JsonDocument.Parse(fence);
                Assert.Equal(GuidanceSurface.RunTool, json.RootElement.GetProperty("tool").GetString());
                Assert.Equal(
                    verb,
                    json.RootElement.GetProperty("arguments").GetProperty("command_name").GetString());
            }
        }

        // view 不是 paragraph 的 schema operation → 不支持提示，无伪 JSON 示例
        var view = InvokeCaptured("help", "docx", "view", "paragraph", "--surface", "agent-json");
        Assert.Equal(0, view.ExitCode);
        Assert.Contains("不支持", view.Stdout, StringComparison.Ordinal);
        Assert.Empty(ExtractJsonFences(view.Stdout));
    }

    [Fact]
    public void AgentHelp_ContainerSetUsesDeclaredOperationInsteadOfReadonlyShortcut()
    {
        var result = InvokeCaptured(
            "help", "docx", "set", "document", "--surface", "agent-json");

        Assert.Equal(0, result.ExitCode);
        Assert.DoesNotContain("只读容器", result.Stdout, StringComparison.Ordinal);
        var fence = Assert.Single(ExtractJsonFences(result.Stdout));
        using var json = JsonDocument.Parse(fence);
        var arguments = json.RootElement.GetProperty("arguments");
        Assert.Equal("set", arguments.GetProperty("command_name").GetString());
        Assert.Contains(
            arguments.GetProperty("command_arguments").EnumerateArray(),
            item => item.GetString() == "author=sample");
    }

    [Fact]
    public void AgentHelp_SetSampleRequiresPropertyToDeclareSet()
    {
        var result = InvokeCaptured(
            "help", "docx", "set", "textbox", "--surface", "agent-json");

        Assert.Equal(0, result.ExitCode);
        var fence = Assert.Single(ExtractJsonFences(result.Stdout));
        using var json = JsonDocument.Parse(fence);
        var commandArguments = json.RootElement.GetProperty("arguments")
            .GetProperty("command_arguments")
            .EnumerateArray()
            .Select(item => item.GetString())
            .ToArray();
        Assert.Contains("width=12pt", commandArguments);
        Assert.DoesNotContain("text=sample", commandArguments);
    }

    [Fact]
    public void AgentHelp_BannerTeachesMutationFastPathAndPartialSuccessBoundary()
    {
        var result = InvokeCaptured("help", "--surface", "agent-json");
        Assert.Equal(0, result.ExitCode);
        var text = result.Stdout;

        Assert.Contains(GuidanceSurface.HelpTool, text, StringComparison.Ordinal);
        Assert.Contains(GuidanceSurface.LoadSkillTool, text, StringComparison.Ordinal);
        Assert.Contains(GuidanceSurface.BatchTool, text, StringComparison.Ordinal);
        Assert.Contains(GuidanceSurface.RunTool, text, StringComparison.Ordinal);
        Assert.Contains("不是事务", text, StringComparison.Ordinal);
        Assert.Contains("不回滚", text, StringComparison.Ordinal);
        Assert.Contains("部分成功", text, StringComparison.Ordinal);
        Assert.Contains("可丢弃副本", text, StringComparison.Ordinal);
        Assert.Contains("逐项", text, StringComparison.Ordinal);
        Assert.Contains("8192", text, StringComparison.Ordinal);
        Assert.Contains("outputFile", text, StringComparison.Ordinal);
        // 明确否定自动草稿/原子覆盖；不得写成已具备这些能力
        Assert.Contains("不会自动创建草稿", text, StringComparison.Ordinal);
        Assert.Contains("不会自动创建草稿、发起审批或原子覆盖", text, StringComparison.Ordinal);
        Assert.DoesNotContain("会自动创建草稿、发起审批并原子覆盖", text, StringComparison.Ordinal);

        var batch = ExtractJsonFences(text)
            .Select(ParseObject)
            .First(o => o.RootElement.GetProperty("tool").GetString() == GuidanceSurface.BatchTool);
        using (batch)
        {
            var ops = batch.RootElement.GetProperty("arguments").GetProperty("operations");
            Assert.True(ops.GetArrayLength() >= 3, "banner Batch must have >= 3 operations");
            foreach (var op in ops.EnumerateArray())
            {
                var cmd = op.GetProperty("command").GetString();
                Assert.Contains(cmd!, SchemaHelpAgentRenderer.BatchMutationCommands);
                Assert.DoesNotContain(cmd!, SchemaHelpAgentRenderer.ReadOrCheckCommands);
            }
        }

        AssertNoForbiddenSyntax(text);
    }

    [Fact]
    public void AgentHelp_BatchExamplesRejectAmbiguousAddMoveShapes()
    {
        // 负向 fixture：多位置字段
        using var ambiguous = JsonDocument.Parse("""
            {"command":"add","parent":"/body","type":"paragraph","index":1,"after":"/body/p[1]"}
            """);
        Assert.True(SchemaHelpAgentRenderer.HasAmbiguousPositionFields(ambiguous.RootElement));

        // 负向 fixture：add.from + props
        using var fromAndProps = JsonDocument.Parse("""
            {"command":"add","parent":"/body","type":"image","from":"/tmp/a.png","props":{"w":"1in"}}
            """);
        Assert.True(SchemaHelpAgentRenderer.HasAddFromAndProps(fromAndProps.RootElement));

        // 合法 banner Batch：位置字段互斥且无 from+props
        var banner = SchemaHelpAgentRenderer.BuildBannerBatchExampleJson();
        using var batch = JsonDocument.Parse(banner);
        foreach (var op in batch.RootElement.GetProperty("arguments").GetProperty("operations").EnumerateArray())
        {
            Assert.False(SchemaHelpAgentRenderer.HasAmbiguousPositionFields(op));
            Assert.False(SchemaHelpAgentRenderer.HasAddFromAndProps(op));
        }
    }

    [Fact]
    public void AgentHelp_ReadCommandsNeverUseBatch()
    {
        // 读取型 Help：元素 get/query 与 banner 中的读取路由说明
        foreach (var verb in new[] { "get", "query" })
        {
            var result = InvokeCaptured(
                "help", "docx", verb, "paragraph", "--surface", "agent-json");
            Assert.Equal(0, result.ExitCode);
            Assert.DoesNotContain(GuidanceSurface.BatchTool, result.Stdout, StringComparison.Ordinal);
            foreach (var fence in ExtractJsonFences(result.Stdout))
            {
                using var json = JsonDocument.Parse(fence);
                Assert.Equal(GuidanceSurface.RunTool, json.RootElement.GetProperty("tool").GetString());
            }
        }

        var banner = InvokeCaptured("help", "--surface", "agent-json");
        Assert.Contains("禁止放入 Batch", banner.Stdout, StringComparison.Ordinal);
        Assert.Contains("8192", banner.Stdout, StringComparison.Ordinal);
        Assert.Contains("outputFile", banner.Stdout, StringComparison.Ordinal);
        Assert.Contains("不归一化", banner.Stdout, StringComparison.Ordinal);
    }

    [Fact]
    public void AgentHelp_UnsafeSchemaExamplesAreNotReused()
    {
        // paragraph schema 含 --prop CLI examples；Agent 输出不得回显这些原串
        using var schema = SchemaHelpLoader.LoadSchema("docx", "paragraph");
        var agent = SchemaHelpAgentRenderer.RenderElement(schema, "set");

        Assert.DoesNotContain("--prop align=center", agent, StringComparison.Ordinal);
        Assert.DoesNotContain("officecli set", agent, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("属性", agent, StringComparison.Ordinal);
        Assert.Contains("路径", agent, StringComparison.Ordinal);
        Assert.Contains(GuidanceSurface.RunTool, agent, StringComparison.Ordinal);
    }

    [Fact]
    public void AgentHelp_RejectsUnknownSurfaceAndJsonModeConflicts()
    {
        var unknown = InvokeCaptured("help", "--surface", "mcp-json");
        Assert.NotEqual(0, unknown.ExitCode);
        Assert.Contains("未知 surface", unknown.Stderr, StringComparison.Ordinal);
        Assert.DoesNotContain("Schema 参考", unknown.Stdout, StringComparison.Ordinal);

        var withJson = InvokeCaptured("help", "docx", "paragraph", "--surface", "agent-json", "--json");
        Assert.NotEqual(0, withJson.ExitCode);
        Assert.Contains("不能与 --json", withJson.Stderr, StringComparison.Ordinal);

        var withJsonl = InvokeCaptured("help", "all", "--surface", "agent-json", "--jsonl");
        Assert.NotEqual(0, withJsonl.ExitCode);
        Assert.Contains("不能与 --json", withJsonl.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public void AgentHelp_UnknownInputsRemainBounded()
    {
        var huge = new string('x', 4096);
        var result = InvokeCaptured("help", "--surface", huge);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("未知 surface", result.Stderr, StringComparison.Ordinal);
        // 错误不得完整回显 4KiB 输入
        Assert.True(result.Stderr.Length < 512, $"stderr too large: {result.Stderr.Length}");
        Assert.DoesNotContain(huge, result.Stderr, StringComparison.Ordinal);

        var unknownFormat = InvokeCaptured("help", huge, "--surface", "agent-json");
        Assert.NotEqual(0, unknownFormat.ExitCode);
        Assert.DoesNotContain(huge, unknownFormat.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public void AgentHelp_AllSchemasRenderDeterministically()
    {
        foreach (var format in SchemaHelpLoader.ListFormats())
        {
            foreach (var element in SchemaHelpLoader.ListElements(format))
            {
                using var schema = SchemaHelpLoader.LoadSchema(format, element);
                var first = SchemaHelpAgentRenderer.RenderElement(schema);
                var second = SchemaHelpAgentRenderer.RenderElement(schema);
                Assert.Equal(first, second);

                foreach (var fence in ExtractJsonFences(first))
                    using (JsonDocument.Parse(fence)) { }
            }
        }
    }

    [Fact]
    public void AgentHelp_AllSchemasExcludeForbiddenSyntax()
    {
        var combined = new System.Text.StringBuilder();
        combined.AppendLine(SchemaHelpAgentRenderer.RenderBanner());
        foreach (var format in SchemaHelpLoader.ListFormats())
        {
            foreach (var element in SchemaHelpLoader.ListElements(format))
            {
                using var schema = SchemaHelpLoader.LoadSchema(format, element);
                combined.AppendLine(SchemaHelpAgentRenderer.RenderElement(schema));
            }
        }

        AssertNoForbiddenSyntax(combined.ToString());
    }

    [Theory]
    [InlineData(null, 0)]
    [InlineData("", 0)]
    [InlineData("cli", 0)]
    [InlineData("agent-json", 1)]
    [InlineData("AGENT-JSON", 1)]
    public void GuidanceSurface_TryParse_AcceptsKnownValues(string? input, int expectedKind)
    {
        Assert.True(GuidanceSurface.TryParse(input, out var surface, out var error));
        Assert.Null(error);
        Assert.Equal((GuidanceSurfaceKind)expectedKind, surface);
    }

    [Fact]
    public void GuidanceSurface_TryParse_RejectsUnknown()
    {
        Assert.False(GuidanceSurface.TryParse("sidecar", out _, out var error));
        Assert.Contains("未知 surface", error, StringComparison.Ordinal);
    }

    private static void AssertNoForbiddenSyntax(string text)
    {
        foreach (var pattern in ForbiddenPatterns)
            Assert.DoesNotContain(pattern, text, StringComparison.OrdinalIgnoreCase);

        // 禁止可执行 officecli <verb>，并禁止任意 officecli / OfficeCLI 字样
        Assert.False(
            OfficeCliCommandRegex.IsMatch(text),
            "Agent Help must not contain executable 'officecli <verb>' commands");
        Assert.DoesNotContain("officecli", text, StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> ExtractJsonFences(string markdown)
    {
        var list = new List<string>();
        foreach (Match match in JsonFenceRegex.Matches(markdown))
            list.Add(match.Groups[1].Value);
        return list;
    }

    private static JsonDocument ParseObject(string json) => JsonDocument.Parse(json);

    private static CommandResult InvokeCaptured(params string[] args)
    {
        lock (typeof(AgentHelpSurfaceTests))
        {
            var oldOut = Console.Out;
            var oldError = Console.Error;
            using var stdout = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
            using var stderr = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
            Console.SetOut(stdout);
            Console.SetError(stderr);
            try
            {
                var root = CommandBuilder.BuildRootCommand();
                var exitCode = root.Parse(args).Invoke();
                return new CommandResult(exitCode, stdout.ToString(), stderr.ToString());
            }
            finally
            {
                Console.SetOut(oldOut);
                Console.SetError(oldError);
            }
        }
    }

    private sealed record CommandResult(int ExitCode, string Stdout, string Stderr);
}
