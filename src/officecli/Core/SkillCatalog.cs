// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Read-only access to embedded Agent guidance. This class never writes into
/// another tool's home directory; callers fetch guidance through load_skill.
/// </summary>
internal static class SkillCatalog
{
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

    public static string BuildSkillTriggerSummary()
    {
        var parts = SkillMap.Keys.Select(name =>
            SkillTriggers.TryGetValue(name, out var trigger) ? $"{name} → {trigger}" : name);
        return "Before you create/add/set/remove on an Office file, run `load_skill <X>` for the relevant workflow. "
            + "Pick X by need: " + string.Join(" · ", parts)
            + ". Run `load_skill` with no name to list all guides.";
    }

    public static string BuildSkillCatalog()
    {
        var sb = new StringBuilder("# officecli skills\n\n");
        sb.Append("Read-only workflow guides bundled with OfficeCLI.\n\n");
        sb.Append("- `load_skill <name>` — the guide and its reference-file manifest\n");
        sb.Append("- `load_skill <name> --path <relpath>` — one text reference file\n\n");
        foreach (var (name, folder) in SkillMap)
        {
            var description = GetSkillDescription(folder);
            sb.Append($"## {name}\n{(description.Length > 0 ? description : "(no description)")}\n\n");
        }
        return sb.ToString().TrimEnd() + "\n";
    }

    public static string KnownSkillsList() => string.Join(", ", SkillMap.Keys.OrderBy(key => key));

    public static string LoadSkillContent(string skillName)
    {
        var folder = ResolveFolder(skillName);
        var content = LoadEmbeddedResource($"skills/{folder}/SKILL.md")
            ?? throw new ArgumentException($"Embedded SKILL.md not found for '{skillName}'");
        return content + BuildReferenceManifest(skillName);
    }

    public static IReadOnlyList<string> ListSkillFiles(string skillName)
    {
        var folder = ResolveFolder(skillName);
        var prefix = $"skills/{folder}/";
        return Assembly.GetExecutingAssembly().GetManifestResourceNames()
            .Where(name => name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Select(name => name[prefix.Length..])
            .Where(path => !path.Equals("SKILL.md", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static string LoadSkillFile(string skillName, string relativePath)
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

    private static string ResolveFolder(string skillName) =>
        SkillMap.TryGetValue(skillName, out var folder)
            ? folder
            : throw new ArgumentException($"Unknown skill: {skillName}. Available: {KnownSkillsList()}");

    private static string GetSkillDescription(string folder)
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

    private static string BuildReferenceManifest(string skillName)
    {
        var files = ListSkillFiles(skillName);
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
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name);
        if (stream is null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
