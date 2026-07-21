// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Author: xiesq
// Created: 2026-07-15

namespace OfficeCli.Core;

/// <summary>
/// Guidance output surface for Help/Skill.
/// Unspecified (null/empty) maps to <see cref="GuidanceSurfaceKind.Cli"/>.
/// Explicit values are only <c>cli</c> and <c>agent-json</c>; unknown values fail-closed
/// and must never silently fall back to CLI.
/// Author: xiesq, 2026-07-15
/// </summary>
internal enum GuidanceSurfaceKind
{
    /// <summary>Default human CLI Help/Skill surface.</summary>
    Cli = 0,

    /// <summary>Host-neutral Agent Tool JSON guidance surface.</summary>
    AgentJson = 1,
}

/// <summary>
/// Strict parser and shared role-placeholder tokens for Agent guidance.
/// OfficeCLI only emits these four host-neutral placeholders; hosts bind real tool names.
/// Author: xiesq, 2026-07-15
/// </summary>
internal static class GuidanceSurface
{
    /// <summary>CLI surface token accepted by <c>--surface</c>.</summary>
    public const string CliToken = "cli";

    /// <summary>Agent JSON surface token accepted by <c>--surface</c>.</summary>
    public const string AgentJsonToken = "agent-json";

    /// <summary>Host-neutral Help tool placeholder.</summary>
    public const string HelpTool = "{{OFFICE_HELP_TOOL}}";

    /// <summary>Host-neutral load_skill tool placeholder.</summary>
    public const string LoadSkillTool = "{{OFFICE_LOAD_SKILL_TOOL}}";

    /// <summary>Host-neutral Batch tool placeholder.</summary>
    public const string BatchTool = "{{OFFICE_BATCH_TOOL}}";

    /// <summary>Host-neutral Run tool placeholder.</summary>
    public const string RunTool = "{{OFFICE_RUN_TOOL}}";

    /// <summary>
    /// Parse a surface token. Empty/null → CLI. Unknown → false with a truncated error.
    /// </summary>
    /// <param name="value">Raw <c>--surface</c> value; null/empty means default CLI.</param>
    /// <param name="surface">Parsed surface when successful.</param>
    /// <param name="error">Stable Chinese error when parsing fails; never echoes full long input.</param>
    /// <returns>True when the value is accepted.</returns>
    public static bool TryParse(string? value, out GuidanceSurfaceKind surface, out string? error)
    {
        // 未传 surface：保持人类 CLI 默认行为
        if (string.IsNullOrEmpty(value))
        {
            surface = GuidanceSurfaceKind.Cli;
            error = null;
            return true;
        }

        // 显式 cli：与默认等价，便于脚本写死参数
        if (string.Equals(value, CliToken, StringComparison.OrdinalIgnoreCase))
        {
            surface = GuidanceSurfaceKind.Cli;
            error = null;
            return true;
        }

        // 显式 agent-json：进入宿主中立 Tool JSON 说明面
        if (string.Equals(value, AgentJsonToken, StringComparison.OrdinalIgnoreCase))
        {
            surface = GuidanceSurfaceKind.AgentJson;
            error = null;
            return true;
        }

        // 未知 surface 必须 fail-closed，禁止静默按 CLI 处理
        surface = default;
        error =
            $"错误：未知 surface “{TruncateForError(value)}”。可用值：{CliToken}、{AgentJsonToken}。";
        return false;
    }

    /// <summary>
    /// Truncate user-controlled tokens in errors so oversized input cannot amplify responses.
    /// </summary>
    internal static string TruncateForError(string value, int maxChars = 64)
    {
        if (string.IsNullOrEmpty(value)) return value;
        if (value.Length <= maxChars) return value;
        // 保留前缀并标注截断，避免把 4KiB 未知输入完整回显到 stderr
        return value.Substring(0, maxChars) + "…";
    }
}
