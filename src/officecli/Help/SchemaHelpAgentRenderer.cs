// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Author: xiesq
// Created: 2026-07-15

using System.Text;
using System.Text.Json;
using OfficeCli.Core;

namespace OfficeCli.Help;

/// <summary>
/// Projects the unique help schema corpus into host-neutral Agent Tool JSON guidance.
/// This renderer only assembles examples from structured schema fields; it never reuses
/// schema CLI/Shell example strings and is not a canonical contract or execution validator.
/// Author: xiesq, 2026-07-15
/// </summary>
internal static class SchemaHelpAgentRenderer
{
    /// <summary>Agent-facing Batch mutation allowlist (phase-0 fast path).</summary>
    internal static readonly HashSet<string> BatchMutationCommands =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "set", "add", "import", "remove", "move", "swap", "raw-set", "add-part",
        };

    /// <summary>Commands that must never appear in Agent Batch examples.</summary>
    internal static readonly HashSet<string> ReadOrCheckCommands =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "get", "query", "view", "raw", "validate",
        };

    // 自由文本中出现这些标记时整段丢弃，不做字符串替换
    private static readonly string[] ForbiddenFreeTextMarkers =
    {
        "officecli",
        "--commands",
        "$FILE",
        "heredoc",
        "| jq",
        "| grep",
        "tool.file.office.",
        "--",
        "```bash",
        "```sh",
        "```shell",
        "```python",
        "```powershell",
        "<<EOF",
        "<<" ,
    };

    // schema 顶层/属性 examples 等 CLI 痕迹标记；命中则整段忽略
    private static readonly string[] CliExampleMarkers =
    {
        "officecli",
        "--prop",
        "--type",
        "--commands",
        "|",
        "$",
        "```",
    };

    private static readonly JsonWriterOptions IndentedWriterOptions =
        new() { Indented = true };

    /// <summary>
    /// No-args Agent Help banner: role placeholders, Batch/Run routing, and one
    /// multi-mutation Batch example (never a length-1 Batch as the single-command path).
    /// </summary>
    internal static string RenderBanner()
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Office Agent Guidance");
        sb.AppendLine();
        sb.AppendLine("本说明面使用宿主中立角色占位符，宿主负责绑定真实 Tool 名称：");
        sb.AppendLine($"- Help：`{GuidanceSurface.HelpTool}`");
        sb.AppendLine($"- Load Skill：`{GuidanceSurface.LoadSkillTool}`");
        sb.AppendLine($"- Batch：`{GuidanceSurface.BatchTool}`");
        sb.AppendLine($"- Run：`{GuidanceSurface.RunTool}`");
        sb.AppendLine();
        sb.AppendLine("## Batch / Run 选择规则");
        sb.AppendLine();
        sb.AppendLine("- 单步、通常 1～2 项 mutation、依赖前序结果、需要丰富诊断，以及全部读取/检查命令（get/query/view/raw/validate/screenshot）一律使用 Run：`command_name` + `command_arguments[]`。");
        sb.AppendLine("- 同一文件通常至少 3 项、参数已知且相互独立的 mutation，才优先使用 Batch。该数量只是 guidance heuristic，不是 schema 硬限制。");
        sb.AppendLine("- Agent Batch 仅承诺 mutation 快路径：`set/add/import/remove/move/swap/raw-set/add-part`。");
        sb.AppendLine("- `get/query/view/raw/validate` 禁止放入 Batch；Run 只禁止 `command_name=batch`，现行允许的单命令仍可走 Run。");
        sb.AppendLine("- 禁止 read-heavy Batch。当前 Batch JSON 超过 8192 bytes 时只返回带 `outputFile` 的精简 envelope，阶段 0 不归一化或搬运完整结果。");
        sb.AppendLine("- `stop_on_error` 只停止后续 operation，不回滚已成功修改；Batch 不是事务。");
        sb.AppendLine("- 只有业务接受部分成功，或调用方已经提供可丢弃副本时才能使用 Batch；必须逐项检查结果，并用独立 Run 完成 readback/view/validate/screenshot。");
        sb.AppendLine("- 全有或全无且没有阶段 0 之外的文件草稿生命周期时，不得对原件使用 Batch。");
        sb.AppendLine("- 阶段 0 不会自动创建草稿、发起审批或原子覆盖。");
        sb.AppendLine("- `add/move` 的 `index/after/before` 最多出现一个；`add.from` 不得与 `props` 同时出现。");
        sb.AppendLine("- `raw-set/add-part` 是需宿主 command approval 的高风险回退；下一步依赖返回 relationship/path 时必须拆成 Run。");
        sb.AppendLine();
        sb.AppendLine("## 多项 mutation 快路径示例");
        sb.AppendLine();
        sb.AppendLine("```json");
        sb.AppendLine(BuildBannerBatchExampleJson());
        sb.AppendLine("```");
        sb.AppendLine();
        sb.AppendLine("## 单命令主路示例（Run）");
        sb.AppendLine();
        sb.AppendLine("```json");
        sb.AppendLine(BuildRunEnvelopeJson(
            "set",
            new[]
            {
                "/workspace/document.docx",
                "/body/p[1]",
                "--prop",
                "bold=true",
            }));
        sb.AppendLine("```");
        sb.AppendLine();
        sb.AppendLine("使用 Help 时传入 format/element（可选 verb）获取结构化字段说明与可执行 Tool JSON 示例。");
        return sb.ToString().TrimEnd('\r', '\n');
    }

    /// <summary>
    /// Project one schema document into Agent Markdown with structure + Run Tool JSON examples.
    /// </summary>
    internal static string RenderElement(JsonDocument doc, string? verbFilter = null)
    {
        var sb = new StringBuilder();
        var root = doc.RootElement;

        var format = root.TryGetProperty("format", out var f) ? f.GetString() ?? "" : "";
        var element = root.TryGetProperty("element", out var e) ? e.GetString() ?? "" : "";
        var isContainer = root.TryGetProperty("container", out var c)
                          && c.ValueKind == JsonValueKind.True;

        var header = verbFilter == null
            ? $"{format} {element}"
            : $"{format} {verbFilter} {element}";
        sb.AppendLine($"# {header}");
        sb.AppendLine();

        // verb 过滤时：不支持则明确提示，不生成伪示例
        if (verbFilter != null
            && root.TryGetProperty("operations", out var opsEl)
            && (!opsEl.TryGetProperty(verbFilter, out var opVal)
                || opVal.ValueKind != JsonValueKind.True))
        {
            sb.AppendLine($"{format} {element} 不支持 '{verbFilter}'。");
            sb.AppendLine();
            sb.AppendLine($"请改用 `{GuidanceSurface.HelpTool}` 查询该元素支持的操作，或使用 `{GuidanceSurface.RunTool}` 调用其他已支持命令。");
            return sb.ToString().TrimEnd('\r', '\n');
        }

        if (isContainer)
            sb.AppendLine("容器元素；可用 mutation 以 operations 与属性声明为准。");

        AppendSafeStringField(sb, root, "description");

        if (root.TryGetProperty("parent", out var parent))
        {
            var parentStr = parent.ValueKind switch
            {
                JsonValueKind.String => parent.GetString() ?? "",
                JsonValueKind.Array => string.Join(", ",
                    parent.EnumerateArray().Select(p => p.GetString() ?? "")),
                _ => "",
            };
            if (!string.IsNullOrEmpty(parentStr))
                sb.AppendLine($"父元素：{parentStr}");
        }

        AppendPaths(sb, root);
        AppendAddressing(sb, root);
        AppendOperations(sb, root, verbFilter);
        AppendProperties(sb, root, verbFilter);
        AppendParts(sb, root);
        AppendChildren(sb, root);

        // note 是自由文本：含 CLI/Shell 标记则整段忽略
        if (root.TryGetProperty("note", out var note)
            && note.ValueKind == JsonValueKind.String
            && note.GetString() is { Length: > 0 } noteStr
            && IsSafeFreeText(noteStr))
        {
            sb.AppendLine();
            sb.AppendLine($"说明：{noteStr}");
        }

        // 可执行示例只从结构化字段生成 Run envelope；不回显 schema examples
        AppendRunExamples(sb, root, format, element, isContainer, verbFilter);

        sb.AppendLine();
        sb.AppendLine("提醒：单命令主路使用 Run。只有同一文件上通常至少 3 项相互独立 mutation，且接受部分成功或已有可丢弃副本时，才优先 Batch；禁止 read-heavy Batch，超过 8192 bytes 的 Batch 结果可能只返回 `outputFile` 精简 envelope。");

        return sb.ToString().TrimEnd('\r', '\n');
    }

    /// <summary>
    /// Agent listing for `help &lt;format&gt;` / verb-filtered element names.
    /// </summary>
    internal static string RenderElementList(
        string format, string? verbFilter, IReadOnlyList<string> elements)
    {
        var sb = new StringBuilder();
        var header = verbFilter == null
            ? $"{format} 的元素（Agent）"
            : $"{format} 中支持“{verbFilter}”的元素（Agent）";
        sb.AppendLine($"# {header}");
        sb.AppendLine();
        if (elements.Count == 0)
        {
            sb.AppendLine(verbFilter == null
                ? "没有可列出的元素。"
                : $"{format} 中没有支持“{verbFilter}”的元素。");
            return sb.ToString().TrimEnd('\r', '\n');
        }

        foreach (var el in elements)
            sb.AppendLine($"- {el}");

        sb.AppendLine();
        sb.AppendLine($"使用 `{GuidanceSurface.HelpTool}` 传入 format/element（可选 verb）查看结构化字段与 Run Tool JSON 示例。");
        sb.AppendLine($"读取/检查与单步 mutation 走 `{GuidanceSurface.RunTool}`；多项独立 mutation 快路径走 `{GuidanceSurface.BatchTool}`。");
        return sb.ToString().TrimEnd('\r', '\n');
    }

    /// <summary>
    /// Project every embedded schema (optionally filtered by format) for Agent `help all`.
    /// </summary>
    internal static string RenderAll(string? onlyFormat = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Office Agent Help — all schemas");
        sb.AppendLine();
        sb.AppendLine(RenderBanner());
        sb.AppendLine();

        var formats = onlyFormat == null
            ? SchemaHelpLoader.ListFormats()
            : new[] { SchemaHelpLoader.NormalizeFormat(onlyFormat) };

        foreach (var format in formats)
        {
            foreach (var element in SchemaHelpLoader.ListElements(format))
            {
                using var doc = SchemaHelpLoader.LoadSchema(format, element);
                sb.AppendLine();
                sb.AppendLine("---");
                sb.AppendLine();
                sb.AppendLine(RenderElement(doc, verbFilter: null));
            }
        }

        return sb.ToString().TrimEnd('\r', '\n');
    }

    /// <summary>
    /// Build a Run tool envelope JSON string with host-neutral tool placeholder.
    /// Uses Utf8JsonWriter so escaping is deterministic and trim-safe.
    /// </summary>
    internal static string BuildRunEnvelopeJson(string commandName, IReadOnlyList<string> commandArguments)
    {
        using var ms = new MemoryStream();
        using (var writer = new Utf8JsonWriter(ms, IndentedWriterOptions))
        {
            writer.WriteStartObject();
            writer.WriteString("tool", GuidanceSurface.RunTool);
            writer.WritePropertyName("arguments");
            writer.WriteStartObject();
            writer.WriteString("command_name", commandName);
            writer.WritePropertyName("command_arguments");
            writer.WriteStartArray();
            foreach (var arg in commandArguments)
                writer.WriteStringValue(arg);
            writer.WriteEndArray();
            writer.WriteEndObject();
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(ms.ToArray());
    }

    /// <summary>
    /// Stable multi-mutation Batch example for the no-args banner.
    /// </summary>
    internal static string BuildBannerBatchExampleJson()
    {
        using var ms = new MemoryStream();
        using (var writer = new Utf8JsonWriter(ms, IndentedWriterOptions))
        {
            writer.WriteStartObject();
            writer.WriteString("tool", GuidanceSurface.BatchTool);
            writer.WritePropertyName("arguments");
            writer.WriteStartObject();
            writer.WriteString("file", "/workspace/report.docx");
            writer.WritePropertyName("operations");
            writer.WriteStartArray();

            WriteSetOperation(writer, "/body/p[1]", "bold", true);
            WriteSetOperation(writer, "/body/p[2]", "italic", true);
            WriteSetOperation(writer, "/body/p[3]", "color", "1F4E79");

            writer.WriteEndArray();
            writer.WriteBoolean("stop_on_error", true);
            writer.WriteEndObject();
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(ms.ToArray());
    }

    /// <summary>
    /// True when free text is safe to surface on the Agent channel.
    /// </summary>
    internal static bool IsSafeFreeText(string text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        foreach (var marker in ForbiddenFreeTextMarkers)
        {
            if (text.Contains(marker, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    /// <summary>
    /// True when a schema example string looks like CLI/Shell and must not be reused.
    /// </summary>
    internal static bool LooksLikeCliOrShellExample(string text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var marker in CliExampleMarkers)
        {
            if (text.Contains(marker, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Validate add/move position fields: at most one of index/after/before.
    /// </summary>
    internal static bool HasAmbiguousPositionFields(JsonElement operation)
    {
        var count = 0;
        if (HasNonNullProperty(operation, "index")) count++;
        if (HasNonNullProperty(operation, "after")) count++;
        if (HasNonNullProperty(operation, "before")) count++;
        return count > 1;
    }

    /// <summary>
    /// Validate add.from + props mutual exclusion.
    /// </summary>
    internal static bool HasAddFromAndProps(JsonElement operation)
    {
        var command = operation.TryGetProperty("command", out var c) ? c.GetString() : null;
        if (!string.Equals(command, "add", StringComparison.OrdinalIgnoreCase))
            return false;
        return HasNonNullProperty(operation, "from") && HasNonNullProperty(operation, "props");
    }

    private static bool HasNonNullProperty(JsonElement obj, string name) =>
        obj.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null;

    private static void WriteSetOperation(
        Utf8JsonWriter writer, string path, string propName, object propValue)
    {
        writer.WriteStartObject();
        writer.WriteString("command", "set");
        writer.WriteString("path", path);
        writer.WritePropertyName("props");
        writer.WriteStartObject();
        switch (propValue)
        {
            case bool b:
                writer.WriteBoolean(propName, b);
                break;
            case int i:
                writer.WriteNumber(propName, i);
                break;
            case long l:
                writer.WriteNumber(propName, l);
                break;
            case double d:
                writer.WriteNumber(propName, d);
                break;
            default:
                writer.WriteString(propName, propValue?.ToString() ?? "");
                break;
        }

        writer.WriteEndObject();
        writer.WriteEndObject();
    }

    private static void AppendSafeStringField(StringBuilder sb, JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var el) || el.ValueKind != JsonValueKind.String)
            return;
        if (el.GetString() is not { Length: > 0 } text) return;
        if (!IsSafeFreeText(text)) return;
        sb.AppendLine(text);
    }

    private static void AppendPaths(StringBuilder sb, JsonElement root)
    {
        if (!root.TryGetProperty("paths", out var paths)) return;
        var pathList = new List<string>();
        if (paths.TryGetProperty("stable", out var stable) && stable.ValueKind == JsonValueKind.Array)
        {
            foreach (var p in stable.EnumerateArray())
                if (p.GetString() is { } s) pathList.Add(s);
        }

        if (paths.TryGetProperty("positional", out var pos) && pos.ValueKind == JsonValueKind.Array)
        {
            foreach (var p in pos.EnumerateArray())
                if (p.GetString() is { } s) pathList.Add(s);
        }

        if (pathList.Count > 0)
            sb.AppendLine($"路径：{string.Join("  ", pathList)}");
    }

    private static void AppendAddressing(StringBuilder sb, JsonElement root)
    {
        if (!root.TryGetProperty("addressing", out var addressing)) return;
        var form = addressing.TryGetProperty("pathForm", out var pf) ? pf.GetString() : null;
        if (!string.IsNullOrEmpty(form))
            sb.AppendLine($"寻址方式：{form}");

        if (addressing.TryGetProperty("key", out var keyEl)
            && keyEl.ValueKind == JsonValueKind.String
            && addressing.TryGetProperty("keyValues", out var kv)
            && kv.ValueKind == JsonValueKind.Array)
        {
            var vals = new List<string>();
            foreach (var v in kv.EnumerateArray())
                if (v.ValueKind == JsonValueKind.String) vals.Add(v.GetString()!);
            if (vals.Count > 0)
                sb.AppendLine($"  {keyEl.GetString()} 可选值：{string.Join(", ", vals)}");
        }
    }

    private static void AppendOperations(StringBuilder sb, JsonElement root, string? verbFilter)
    {
        if (!root.TryGetProperty("operations", out var ops) || ops.ValueKind != JsonValueKind.Object)
            return;

        var active = new List<string>();
        foreach (var op in ops.EnumerateObject())
        {
            if (op.Value.ValueKind != JsonValueKind.True) continue;
            if (verbFilter != null
                && !string.Equals(op.Name, verbFilter, StringComparison.OrdinalIgnoreCase))
                continue;
            active.Add(op.Name);
        }

        if (active.Count > 0)
            sb.AppendLine($"操作：{string.Join(" ", active)}");
    }

    private static void AppendProperties(
        StringBuilder sb, JsonElement root, string? verbFilter)
    {
        if (!root.TryGetProperty("properties", out var props)
            || props.ValueKind != JsonValueKind.Object
            || !props.EnumerateObject().Any())
            return;

        sb.AppendLine();
        sb.AppendLine(verbFilter == null ? "属性：" : $"属性（{verbFilter}）：");
        var shown = 0;
        foreach (var prop in props.EnumerateObject())
        {
            if (verbFilter != null)
            {
                if (!prop.Value.TryGetProperty(verbFilter, out var pv)
                    || pv.ValueKind != JsonValueKind.True)
                    continue;
            }

            RenderProperty(sb, prop);
            shown++;
        }

        if (verbFilter != null && shown == 0)
            sb.AppendLine($"  （此元素没有参与 '{verbFilter}' 的属性）");
    }

    private static void RenderProperty(StringBuilder sb, JsonProperty prop)
    {
        var name = prop.Name;
        var body = prop.Value;
        var type = body.TryGetProperty("type", out var t) ? t.GetString() ?? "" : "";

        var opList = new List<string>();
        foreach (var op in new[] { "add", "set", "get" })
        {
            if (body.TryGetProperty(op, out var val) && val.ValueKind == JsonValueKind.True)
                opList.Add(op);
        }

        var opsStr = opList.Count > 0 ? string.Join("/", opList) : "-";
        sb.AppendLine($"  {name}   {type}   [{opsStr}]");

        if (body.TryGetProperty("description", out var desc)
            && desc.GetString() is { Length: > 0 } dstr
            && IsSafeFreeText(dstr))
            sb.AppendLine($"    说明：{dstr}");

        if (body.TryGetProperty("values", out var values) && values.ValueKind == JsonValueKind.Array)
        {
            var vlist = values.EnumerateArray()
                .Select(v => v.GetString())
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();
            if (vlist.Count > 0)
                sb.AppendLine($"    可选值：{string.Join(", ", vlist!)}");
        }

        // 故意不回显 property.examples（通常含 --prop CLI 片段）
        if (body.TryGetProperty("readback", out var rb)
            && rb.GetString() is { Length: > 0 } rbstr
            && IsSafeFreeText(rbstr))
            sb.AppendLine($"    读取结果：{rbstr}");
    }

    private static void AppendParts(StringBuilder sb, JsonElement root)
    {
        if (!root.TryGetProperty("parts", out var parts)
            || parts.ValueKind != JsonValueKind.Array
            || parts.GetArrayLength() == 0)
            return;

        sb.AppendLine();
        sb.AppendLine("Part：");
        foreach (var pt in parts.EnumerateArray())
        {
            var name = pt.TryGetProperty("name", out var nm) ? nm.GetString() ?? "" : "";
            var desc = pt.TryGetProperty("desc", out var ds) ? ds.GetString() ?? "" : "";
            if (!string.IsNullOrEmpty(desc) && !IsSafeFreeText(desc))
                desc = "";
            sb.AppendLine($"  {name}  {desc}".TrimEnd());
        }
    }

    private static void AppendChildren(StringBuilder sb, JsonElement root)
    {
        if (!root.TryGetProperty("children", out var children)
            || children.ValueKind != JsonValueKind.Array
            || children.GetArrayLength() == 0)
            return;

        sb.AppendLine();
        sb.AppendLine("子元素：");
        foreach (var child in children.EnumerateArray())
        {
            var el = child.TryGetProperty("element", out var ce) ? ce.GetString() : "?";
            var seg = child.TryGetProperty("pathSegment", out var ps) ? ps.GetString() : "?";
            var card = child.TryGetProperty("cardinality", out var cd) ? cd.GetString() : "?";
            sb.AppendLine($"  {el}  ({card})  /{seg}");
        }
    }

    private static void AppendRunExamples(
        StringBuilder sb,
        JsonElement root,
        string format,
        string element,
        bool isContainer,
        string? verbFilter)
    {
        if (!root.TryGetProperty("operations", out var ops) || ops.ValueKind != JsonValueKind.Object)
            return;

        var fileToken = FormatWorkspaceFile(format);
        var firstPath = PickFirstPath(root) ?? $"/{element}";
        var parentPath = DeriveParentPath(firstPath);

        var addParents = new List<string>();
        if (root.TryGetProperty("addParent", out var apEl))
        {
            if (apEl.ValueKind == JsonValueKind.String && apEl.GetString() is { } aps)
                addParents.Add(aps);
            else if (apEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var p in apEl.EnumerateArray())
                    if (p.GetString() is { } ps) addParents.Add(ps);
            }
        }

        if (addParents.Count == 0)
            addParents.Add(parentPath);

        bool Has(string v) =>
            ops.TryGetProperty(v, out var ov) && ov.ValueKind == JsonValueKind.True;

        bool Want(string v) =>
            verbFilter == null
            || string.Equals(verbFilter, v, StringComparison.OrdinalIgnoreCase);

        sb.AppendLine();
        sb.AppendLine("## 可执行示例（Run）");
        sb.AppendLine();
        sb.AppendLine("单命令主路始终使用 Run envelope；不要用长度为 1 的 Batch 替代。");

        void Emit(string commandName, IReadOnlyList<string> args)
        {
            sb.AppendLine();
            sb.AppendLine("```json");
            sb.AppendLine(BuildRunEnvelopeJson(commandName, args));
            sb.AppendLine("```");
        }

        // 只为 schema 声明为 true 的 operation 生成示例
        if (Has("add") && !isContainer && Want("add"))
        {
            var args = new List<string> { fileToken, addParents[0], "--type", element };
            var sampleProp = PickSamplePropToken(root, "add");
            if (sampleProp != null)
            {
                args.Add("--prop");
                args.Add(sampleProp);
            }

            Emit("add", args);
        }

        if (Has("set") && Want("set"))
        {
            var sampleProp = PickSamplePropToken(root, "set");
            if (sampleProp != null)
            {
                var args = new List<string> { fileToken, firstPath, "--prop", sampleProp };
                Emit("set", args);
            }
        }

        if (Has("get") && Want("get"))
            Emit("get", new[] { fileToken, firstPath });

        if (Has("query") && Want("query"))
            Emit("query", new[] { fileToken, element });

        if (Has("remove") && !isContainer && Want("remove"))
            Emit("remove", new[] { fileToken, firstPath });

        // view/raw/validate 通常不是 element operations；若 schema 意外声明也走 Run
        foreach (var readCmd in new[] { "view", "raw", "validate" })
        {
            if (Has(readCmd) && Want(readCmd))
                Emit(readCmd, new[] { fileToken });
        }
    }

    private static string FormatWorkspaceFile(string format) =>
        format.ToLowerInvariant() switch
        {
            "xlsx" => "/workspace/workbook.xlsx",
            "pptx" => "/workspace/deck.pptx",
            _ => "/workspace/document.docx",
        };

    private static string? PickFirstPath(JsonElement root)
    {
        if (!root.TryGetProperty("paths", out var paths)) return null;
        if (paths.TryGetProperty("positional", out var pos)
            && pos.ValueKind == JsonValueKind.Array
            && pos.GetArrayLength() > 0
            && pos[0].GetString() is { Length: > 0 } p0)
            return p0;
        if (paths.TryGetProperty("stable", out var stable)
            && stable.ValueKind == JsonValueKind.Array
            && stable.GetArrayLength() > 0
            && stable[0].GetString() is { Length: > 0 } s0)
            return s0;
        return null;
    }

    private static string DeriveParentPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return path;
        var trimmed = path.TrimEnd('/');
        var lastSlash = trimmed.LastIndexOf('/');
        if (lastSlash < 0) return path;
        if (lastSlash == 0) return "/";
        return trimmed.Substring(0, lastSlash);
    }

    /// <summary>
    /// Pick a stable scalar sample prop token from schema property metadata (never CLI examples).
    /// </summary>
    private static string? PickSamplePropToken(JsonElement root, string? verbFilter)
    {
        if (!root.TryGetProperty("properties", out var props) || props.ValueKind != JsonValueKind.Object)
            return null;

        foreach (var prop in props.EnumerateObject())
        {
            if (verbFilter != null
                && (!prop.Value.TryGetProperty(verbFilter, out var flag)
                    || flag.ValueKind != JsonValueKind.True))
                continue;

            var type = prop.Value.TryGetProperty("type", out var t) ? t.GetString() ?? "string" : "string";
            var sample = StableSampleValue(type, prop.Value);
            if (sample == null) continue;
            return $"{prop.Name}={sample}";
        }

        return null;
    }

    private static string? StableSampleValue(string type, JsonElement propBody)
    {
        // enum：取第一个合法 values 项
        if (propBody.TryGetProperty("values", out var values)
            && values.ValueKind == JsonValueKind.Array
            && values.GetArrayLength() > 0
            && values[0].GetString() is { Length: > 0 } first)
            return first;

        return type.ToLowerInvariant() switch
        {
            "boolean" or "bool" => "true",
            "number" or "integer" or "int" => "1",
            "length" => "12pt",
            "color" => "1F4E79",
            _ => "sample",
        };
    }
}
