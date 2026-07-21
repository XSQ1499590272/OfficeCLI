// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Author: xiesq
// Created: 2026-07-15

using System.Collections.Immutable;
using System.Reflection;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Read-only access to embedded Agent guidance.
/// Default (no-surface) methods keep reading <c>skills/</c> for CLI/MCP.
/// <c>agent-json</c> reads only <c>agent-skills/</c> via per-skill manifests and never falls back.
/// Author: xiesq, 2026-07-15
/// </summary>
internal static class SkillCatalog
{
    /// <summary>External skill name → embedded folder under skills/ or agent-skills/.</summary>
    private static readonly Dictionary<string, string> SkillMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["pptx"] = "officecli-pptx",
        ["word"] = "officecli-docx",
        ["excel"] = "officecli-xlsx",
        ["morph-ppt"] = "morph-ppt",
        ["morph-ppt-3d"] = "morph-ppt-3d",
        ["pitch-deck"] = "officecli-pitch-deck",
        ["academic-paper"] = "officecli-academic-paper",
        ["data-dashboard"] = "officecli-data-dashboard",
        ["financial-model"] = "officecli-financial-model",
        ["word-form"] = "officecli-word-form",
    };

    private static readonly Dictionary<string, string> SkillTriggers = new(StringComparer.OrdinalIgnoreCase)
    {
        ["pptx"] = "slide decks / presentations",
        ["word"] = "Word docs, reports, letters, memos",
        ["excel"] = "spreadsheets, financial models, dashboards",
        ["word-form"] = "fillable forms, content controls, protected docs",
        ["morph-ppt"] = "cross-slide Morph animation / continuous motion",
        ["morph-ppt-3d"] = "3D Morph decks (GLB models, camera)",
        ["pitch-deck"] = "fundraising / investor decks (seed, Series A/B/C)",
        ["academic-paper"] = "academic papers / research reports",
        ["data-dashboard"] = "data dashboards",
        ["financial-model"] = "financial models / projections",
    };

    private static readonly HashSet<string> BinarySkillExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pptx", ".docx", ".xlsx", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".glb", ".pdf", ".zip", ".ico",
    };

    // Agent --path 额外拒绝脚本扩展，防止把可执行 recipe 当 reference
    private static readonly HashSet<string> AgentBlockedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pptx", ".docx", ".xlsx", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".glb", ".pdf", ".zip", ".ico",
        ".sh", ".py", ".ps1",
    };

    // Agent 文本资源上限（UTF-8 bytes）
    private const int AgentTextResourceMaxBytes = 128 * 1024;

    private static readonly object AgentIndexLock = new();
    private static ImmutableDictionary<string, AgentSkillEntry>? _agentIndex;
    private static ImmutableDictionary<string, string> _agentIndexErrors =
        ImmutableDictionary<string, string>.Empty.WithComparers(StringComparer.OrdinalIgnoreCase);

    /// <summary>Immutable Agent skill entry built from a valid manifest + SKILL.md.</summary>
    internal sealed record AgentSkillEntry(
        string Name,
        string Folder,
        string Description,
        string Trigger,
        ImmutableArray<string> References);

    internal sealed class AgentManifestException(string code, string message)
        : InvalidOperationException(message)
    {
        internal string Code { get; } = code;
    }

    // ── CLI 默认入口（MCP 继续调用这些无 surface overload）────────────────

    /// <summary>CLI/MCP catalog — always the default <c>skills/</c> surface.</summary>
    public static string BuildSkillCatalog() => BuildSkillCatalog(GuidanceSurfaceKind.Cli);

    /// <summary>CLI/MCP main skill content.</summary>
    public static string LoadSkillContent(string skillName) =>
        LoadSkillContent(skillName, GuidanceSurfaceKind.Cli);

    /// <summary>CLI/MCP reference file list.</summary>
    public static IReadOnlyList<string> ListSkillFiles(string skillName) =>
        ListSkillFiles(skillName, GuidanceSurfaceKind.Cli);

    /// <summary>CLI/MCP reference file content.</summary>
    public static string LoadSkillFile(string skillName, string relativePath) =>
        LoadSkillFile(skillName, relativePath, GuidanceSurfaceKind.Cli);

    public static string BuildSkillTriggerSummary()
    {
        var parts = SkillMap.Keys.Select(name =>
            SkillTriggers.TryGetValue(name, out var trigger) ? $"{name} → {trigger}" : name);
        return "Before you create/add/set/remove on an Office file, run `load_skill <X>` for the relevant workflow. "
            + "Pick X by need: " + string.Join(" · ", parts)
            + ". Run `load_skill` with no name to list all guides.";
    }

    public static string KnownSkillsList() => string.Join(", ", SkillMap.Keys.OrderBy(key => key));

    /// <summary>Stable SkillMap order (insertion order of the dictionary initializer).</summary>
    internal static IReadOnlyList<string> SkillMapNamesInOrder() => SkillMap.Keys.ToList();

    /// <summary>Resolve external skill name to folder; throws if unknown.</summary>
    internal static string ResolveFolder(string skillName) =>
        SkillMap.TryGetValue(skillName, out var folder)
            ? folder
            : throw new ArgumentException($"Unknown skill: {skillName}. Available: {KnownSkillsList()}");

    // ── Surface-aware overload ───────────────────────────────────────────

    /// <summary>
    /// Build catalog for the given surface.
    /// Agent catalog only lists skills with a complete agent-skills manifest + SKILL.md.
    /// </summary>
    public static string BuildSkillCatalog(GuidanceSurfaceKind surface)
    {
        if (surface == GuidanceSurfaceKind.Cli)
        {
            var sb = new StringBuilder("# officecli skills\n\n");
            sb.Append("Read-only workflow guides bundled with OfficeCLI.\n\n");
            sb.Append("- `load_skill <name>` — the guide and its reference-file manifest\n");
            sb.Append("- `load_skill <name> --path <relpath>` — one text reference file\n\n");
            foreach (var (name, folder) in SkillMap)
            {
                var description = GetCliSkillDescription(folder);
                sb.Append($"## {name}\n{(description.Length > 0 ? description : "(no description)")}\n\n");
            }
            return sb.ToString().TrimEnd() + "\n";
        }

        // Agent catalog：只列出完整 Agent 资源；顺序遵循既有 SkillMap。
        // Agent 面向模型的文案禁止出现 officecli / OfficeCLI 字样，避免诱导向命令行。
        var agentSb = new StringBuilder("# Agent skills\n\n");
        agentSb.Append("Host-neutral Tool JSON workflow guides. Use role placeholders only.\n\n");
        agentSb.Append($"- `{GuidanceSurface.LoadSkillTool}` — load a skill or reference\n");
        agentSb.Append($"- `{GuidanceSurface.HelpTool}` — element/property details\n\n");
        foreach (var name in SkillMap.Keys)
        {
            if (!TryGetAgentEntry(name, out var entry) || entry is null) continue;
            agentSb.Append($"## {entry.Name}\n");
            agentSb.Append(entry.Description).Append('\n');
            if (!string.IsNullOrEmpty(entry.Trigger))
                agentSb.Append("Trigger: ").Append(entry.Trigger).Append('\n');
            agentSb.Append('\n');
        }

        return agentSb.ToString().TrimEnd() + "\n";
    }

    /// <summary>Load main skill content for the given surface.</summary>
    public static string LoadSkillContent(string skillName, GuidanceSurfaceKind surface)
    {
        if (surface == GuidanceSurfaceKind.Cli)
        {
            var folder = ResolveFolder(skillName);
            var content = LoadEmbeddedResource($"skills/{folder}/SKILL.md")
                ?? throw new ArgumentException($"Embedded SKILL.md not found for '{skillName}'");
            return content + BuildCliReferenceManifest(skillName);
        }

        var entry = RequireAgentEntry(skillName);
        var agentContent = LoadAgentResource(entry.Folder, "SKILL.md")
            ?? throw new ArgumentException($"Agent SKILL.md not found for '{skillName}'");
        EnforceAgentTextSize(agentContent, skillName, "SKILL.md");
        return agentContent + BuildAgentReferenceManifest(entry);
    }

    /// <summary>List reference paths for the given surface.</summary>
    public static IReadOnlyList<string> ListSkillFiles(string skillName, GuidanceSurfaceKind surface)
    {
        if (surface == GuidanceSurfaceKind.Cli)
        {
            var folder = ResolveFolder(skillName);
            var prefix = $"skills/{folder}/";
            return Assembly.GetExecutingAssembly().GetManifestResourceNames()
                .Where(name => name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                .Select(name => name[prefix.Length..].Replace('\\', '/'))
                .Where(path => !path.Equals("SKILL.md", StringComparison.OrdinalIgnoreCase))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        return RequireAgentEntry(skillName).References;
    }

    /// <summary>
    /// Load one reference file. Agent surface requires an exact manifest entry and never falls back to CLI skills.
    /// </summary>
    public static string LoadSkillFile(
        string skillName, string relativePath, GuidanceSurfaceKind surface)
    {
        if (surface == GuidanceSurfaceKind.Cli)
            return LoadCliSkillFile(skillName, relativePath);

        var entry = RequireAgentEntry(skillName);
        var path = NormalizeAgentRelativePath(relativePath);

        // 精确 allowlist：未声明即使嵌入也不可读
        if (!entry.References.Contains(path, StringComparer.Ordinal))
            throw new ArgumentException(
                $"Agent skill reference not declared in manifest: {GuidanceSurface.TruncateForError(path)}. "
                + $"Use {GuidanceSurface.LoadSkillTool} without --path to see allowed references.");

        var content = LoadAgentResource(entry.Folder, path)
            ?? throw new ArgumentException(
                $"Agent skill reference missing from embedded resources: {GuidanceSurface.TruncateForError(path)}");
        EnforceAgentTextSize(content, skillName, path);
        return content;
    }

    /// <summary>Names present in the Agent catalog (complete resources only), SkillMap order.</summary>
    internal static IReadOnlyList<string> ListAgentSkillNames()
    {
        var names = new List<string>();
        foreach (var name in SkillMap.Keys)
        {
            if (TryGetAgentEntry(name, out _))
                names.Add(name);
        }

        return names;
    }

    // ── Agent index ──────────────────────────────────────────────────────

    private static AgentSkillEntry RequireAgentEntry(string skillName)
    {
        if (!SkillMap.ContainsKey(skillName))
            throw new ArgumentException($"Unknown skill: {skillName}. Available: {KnownSkillsList()}");
        if (!TryGetAgentEntry(skillName, out var entry) || entry is null)
        {
            var reason = _agentIndexErrors.TryGetValue(skillName, out var category)
                ? $"invalid manifest category: {category}"
                : "missing manifest/SKILL.md";
            throw new ArgumentException(
                $"Agent skill '{skillName}' is not available on agent-json surface "
                + $"({reason}).");
        }
        return entry;
    }

    private static bool TryGetAgentEntry(string skillName, out AgentSkillEntry? entry)
    {
        entry = null;
        var index = GetAgentIndex();
        if (!SkillMap.TryGetValue(skillName, out _)) return false;
        // 用 SkillMap 的大小写不敏感键查找
        foreach (var (name, value) in index)
        {
            if (string.Equals(name, skillName, StringComparison.OrdinalIgnoreCase))
            {
                entry = value;
                return true;
            }
        }

        return false;
    }

    private static ImmutableDictionary<string, AgentSkillEntry> GetAgentIndex()
    {
        if (_agentIndex != null) return _agentIndex;
        lock (AgentIndexLock)
        {
            if (_agentIndex != null) return _agentIndex;
            _agentIndex = BuildAgentIndex(out _agentIndexErrors);
            return _agentIndex;
        }
    }

    /// <summary>
    /// Scan embedded agent-skills/*/manifest.json. Invalid entries are skipped (fail-closed per skill).
    /// </summary>
    private static ImmutableDictionary<string, AgentSkillEntry> BuildAgentIndex(
        out ImmutableDictionary<string, string> errors)
    {
        var builder = ImmutableDictionary.CreateBuilder<string, AgentSkillEntry>(
            StringComparer.OrdinalIgnoreCase);
        var errorBuilder = ImmutableDictionary.CreateBuilder<string, string>(
            StringComparer.OrdinalIgnoreCase);
        var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (externalName, folder) in SkillMap)
        {
            try
            {
                var manifestJson = LoadAgentResource(folder, "manifest.json");
                if (manifestJson is null) continue;

                var entry = ParseAgentManifest(
                    externalName,
                    folder,
                    manifestJson,
                    path => LoadAgentResource(folder, path) is not null);

                if (!seenNames.Add(entry.Name))
                    throw new AgentManifestException("duplicate-name", "duplicate manifest name");

                builder[externalName] = entry;
            }
            catch (AgentManifestException ex)
            {
                errorBuilder[externalName] = ex.Code;
            }
            catch (JsonException)
            {
                errorBuilder[externalName] = "invalid-json";
            }
            catch (Exception)
            {
                errorBuilder[externalName] = "invalid-resource";
            }
        }

        errors = errorBuilder.ToImmutable();
        return builder.ToImmutable();
    }

    /// <summary>
    /// Parse and validate one fixed-shape Agent manifest. Resource existence is injected so
    /// the same fail-closed parser can be exercised with extreme fixtures.
    /// </summary>
    internal static AgentSkillEntry ParseAgentManifest(
        string externalName,
        string folder,
        string manifestJson,
        Func<string, bool> resourceExists)
    {
        using var doc = JsonDocument.Parse(manifestJson);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new AgentManifestException("root-not-object", "manifest root must be object");

        var allowed = new HashSet<string>(StringComparer.Ordinal)
            { "name", "description", "trigger", "references" };
        var seenFields = new HashSet<string>(StringComparer.Ordinal);
        foreach (var prop in root.EnumerateObject())
        {
            if (!allowed.Contains(prop.Name))
                throw new AgentManifestException("unknown-field", "manifest contains unknown field");
            if (!seenFields.Add(prop.Name))
                throw new AgentManifestException("duplicate-field", "manifest contains duplicate field");
        }

        var manifestName = RequireManifestString(root, "name");
        var description = RequireManifestString(root, "description");
        var trigger = RequireManifestString(root, "trigger");
        if (!string.Equals(manifestName, externalName, StringComparison.Ordinal))
            throw new AgentManifestException("name-mismatch", "manifest name does not match skill map");

        var references = ParseReferences(root, folder);
        if (!resourceExists("SKILL.md"))
            throw new AgentManifestException("missing-skill", "manifest SKILL.md is missing");
        foreach (var reference in references)
        {
            if (!resourceExists(reference))
                throw new AgentManifestException("missing-reference", "manifest reference is missing");
        }

        return new AgentSkillEntry(
            manifestName, folder, description, trigger, references);
    }

    private static string RequireManifestString(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(value.GetString()))
            throw new AgentManifestException($"missing-{name}", $"manifest {name} must be a non-empty string");
        return value.GetString()!;
    }

    private static ImmutableArray<string> ParseReferences(JsonElement root, string folder)
    {
        if (!root.TryGetProperty("references", out var refsEl))
            throw new AgentManifestException("missing-references", "manifest references are required");
        if (refsEl.ValueKind != JsonValueKind.Array)
            throw new AgentManifestException("references-type", "manifest references must be an array");

        var list = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in refsEl.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String || item.GetString() is not { } raw)
                throw new AgentManifestException("reference-type", "manifest reference must be string");
            var path = NormalizeAgentRelativePath(raw);
            if (!seen.Add(path))
                throw new AgentManifestException("duplicate-reference", "manifest contains duplicate reference");
            // 禁止把主文件放进 references
            if (path.Equals("SKILL.md", StringComparison.OrdinalIgnoreCase)
                || path.Equals("manifest.json", StringComparison.OrdinalIgnoreCase))
                throw new AgentManifestException("reserved-reference", "manifest contains reserved reference");
            if (Path.GetExtension(path) is not ".md" and not ".txt")
                throw new AgentManifestException("reference-extension", "manifest reference must be .md or .txt");
            list.Add(path);
        }

        return list.ToImmutableArray();
    }

    /// <summary>
    /// Normalize and validate an Agent relative path: forward slashes only, no traversal, no binary.
    /// Backslashes are rejected on the raw input before any slash rewriting (O-R03).
    /// Author: xiesq, 2026-07-15
    /// </summary>
    internal static string NormalizeAgentRelativePath(string? relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
            throw new ArgumentException(
                "path is empty — pass a relative skill file declared in the Agent manifest");

        // 必须在 Replace/规范化之前拒绝反斜杠，否则 allowlist 别名可被绕过
        if (relativePath.Contains('\\', StringComparison.Ordinal))
            throw new ArgumentException("Agent skill paths must use forward slashes");

        var path = relativePath.TrimStart('/');
        if (path.Length == 0)
            throw new ArgumentException("path is empty");
        if (Path.IsPathRooted(relativePath) || path.StartsWith('/'))
            throw new ArgumentException("Agent skill paths must be relative");
        if (path.Split('/').Any(segment => segment is ".." or "." or ""))
            throw new ArgumentException(
                $"Invalid Agent skill file path: {GuidanceSurface.TruncateForError(relativePath)}");
        if (AgentBlockedExtensions.Contains(Path.GetExtension(path)))
            throw new ArgumentException(
                $"'{GuidanceSurface.TruncateForError(path)}' is not an Agent-safe text reference.");
        return path;
    }

    private static string? LoadAgentResource(string folder, string relativePath)
    {
        var logical = $"agent-skills/{folder}/{relativePath.Replace('\\', '/')}";
        return LoadEmbeddedResource(logical);
    }

    private static void EnforceAgentTextSize(string content, string skillName, string path)
    {
        var bytes = Encoding.UTF8.GetByteCount(content);
        if (bytes > AgentTextResourceMaxBytes)
            throw new ArgumentException(
                $"Agent skill resource exceeds {AgentTextResourceMaxBytes} bytes: "
                + $"{skillName}/{GuidanceSurface.TruncateForError(path)}");
    }

    private static string BuildAgentReferenceManifest(AgentSkillEntry entry)
    {
        if (entry.References.Length == 0) return "";
        var sb = new StringBuilder("\n\n## Reference files (Agent manifest allowlist)\n\n");
        sb.Append($"Fetch with `{GuidanceSurface.LoadSkillTool}` and a declared reference path argument.\n\n");
        foreach (var file in entry.References)
            sb.Append("- `").Append(file).Append("`\n");
        return sb.ToString();
    }

    // ── CLI helpers (unchanged semantics) ────────────────────────────────

    private static string LoadCliSkillFile(string skillName, string relativePath)
    {
        var folder = ResolveFolder(skillName);
        var path = (relativePath ?? "").Replace('\\', '/').TrimStart('/');
        if (path.Length == 0)
            throw new ArgumentException("path is empty — pass a relative skill file, e.g. reference/decision-rules.md");
        if (path.Split('/').Any(segment => segment is ".." or "."))
            throw new ArgumentException($"Invalid skill file path: {relativePath}");
        if (BinarySkillExtensions.Contains(Path.GetExtension(path)))
            throw new ArgumentException($"'{path}' is a binary asset and cannot be served over the text channel.");

        return LoadEmbeddedResource($"skills/{folder}/{path}")
            ?? throw new ArgumentException(
                $"Skill file not found: {path}. List available files via: officecli load_skill {skillName}");
    }

    private static string GetCliSkillDescription(string folder)
    {
        var content = LoadEmbeddedResource($"skills/{folder}/SKILL.md");
        if (content is null || !content.StartsWith("---", StringComparison.Ordinal)) return "";
        var end = content.IndexOf("---", 3, StringComparison.Ordinal);
        if (end < 0) return "";
        foreach (var line in content[3..end].Split('\n'))
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith("description:", StringComparison.OrdinalIgnoreCase))
                return trimmed["description:".Length..].Trim().Trim('"');
        }
        return "";
    }

    private static string BuildCliReferenceManifest(string skillName)
    {
        var files = ListSkillFiles(skillName, GuidanceSurfaceKind.Cli);
        if (files.Count == 0) return "";

        var shallow = new List<string>();
        var deepGroups = new SortedDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var file in files)
        {
            var segments = file.Split('/');
            if (segments.Length <= 2 || segments[^1].Equals("INDEX.md", StringComparison.OrdinalIgnoreCase))
                shallow.Add(file);
            else
            {
                var group = segments[0] + "/" + segments[1] + "/";
                deepGroups[group] = deepGroups.GetValueOrDefault(group) + 1;
            }
        }

        var sb = new StringBuilder("\n\n## Reference files (bundled with this skill)\n\n");
        sb.Append("Fetch text files with `load_skill ").Append(skillName).Append(" --path <relpath>`. ");
        sb.Append("Binary assets are not exposed through this text interface.\n\n");
        foreach (var file in shallow) sb.Append("- `").Append(file).Append("`\n");
        foreach (var (group, count) in deepGroups)
            sb.Append("- `").Append(group).Append("` — ").Append(count).Append(" files\n");
        return sb.ToString();
    }

    private static string? LoadEmbeddedResource(string name)
    {
        // 同时尝试正斜杠逻辑名（csproj 已归一化）
        var asm = Assembly.GetExecutingAssembly();
        using var stream = asm.GetManifestResourceStream(name)
            ?? asm.GetManifestResourceStream(name.Replace('/', '\\'));
        if (stream is null) return null;
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }
}
