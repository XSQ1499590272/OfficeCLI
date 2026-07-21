// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Author: xiesq
// Created: 2026-07-15

namespace OfficeCli.Core;

/// <summary>
/// Manual CLI parsing for <c>load_skill</c> (early-dispatch, not System.CommandLine).
/// Accepts <c>--surface</c> / <c>--path</c> before or after the skill name; rejects
/// <c>--surface=...</c>, duplicates, missing values, unknown options, and extra positionals.
/// Author: xiesq, 2026-07-15
/// </summary>
internal static class LoadSkillCli
{
    /// <summary>
    /// Dispatch <c>load_skill</c> tokens (everything after the verb).
    /// Returns process exit code.
    /// </summary>
    public static int Dispatch(IReadOnlyList<string> tokens, TextWriter stdout, TextWriter stderr)
    {
        string? skillRelPath = null;
        string? surfaceRaw = null;
        var pathSeen = false;
        var surfaceSeen = false;
        var positional = new List<string>();

        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (token.StartsWith("--surface=", StringComparison.Ordinal)
                || token.StartsWith("--path=", StringComparison.Ordinal))
            {
                stderr.WriteLine(
                    $"错误：不支持 `{token.Split('=')[0]}=...` 写法，请使用 `{token.Split('=')[0]} <value>`。");
                return 1;
            }

            if (token == "--path")
            {
                if (pathSeen)
                {
                    stderr.WriteLine("错误：--path 只能出现一次。");
                    return 1;
                }

                pathSeen = true;
                if (i + 1 >= tokens.Count
                    || string.IsNullOrWhiteSpace(tokens[i + 1])
                    || tokens[i + 1].StartsWith('-'))
                {
                    stderr.WriteLine("错误：--path 需要一个相对路径值。");
                    return 1;
                }

                skillRelPath = tokens[++i];
                continue;
            }

            if (token == "--surface")
            {
                if (surfaceSeen)
                {
                    stderr.WriteLine("错误：--surface 只能出现一次。");
                    return 1;
                }

                surfaceSeen = true;
                if (i + 1 >= tokens.Count
                    || string.IsNullOrWhiteSpace(tokens[i + 1])
                    || tokens[i + 1].StartsWith('-'))
                {
                    stderr.WriteLine("错误：--surface 需要一个值（cli 或 agent-json）。");
                    return 1;
                }

                surfaceRaw = tokens[++i];
                continue;
            }

            if (token.StartsWith('-'))
            {
                stderr.WriteLine(
                    $"错误：未知选项 `{GuidanceSurface.TruncateForError(token)}`。");
                return 1;
            }

            positional.Add(token);
        }

        if (!GuidanceSurface.TryParse(surfaceRaw, out var surface, out var surfaceError))
        {
            stderr.WriteLine(surfaceError);
            return 1;
        }

        if (positional.Count == 0 && string.IsNullOrEmpty(skillRelPath))
        {
            stdout.Write(SkillCatalog.BuildSkillCatalog(surface));
            return 0;
        }

        if (positional.Count == 1)
        {
            try
            {
                stdout.Write(string.IsNullOrEmpty(skillRelPath)
                    ? SkillCatalog.LoadSkillContent(positional[0], surface)
                    : SkillCatalog.LoadSkillFile(positional[0], skillRelPath!, surface));
                return 0;
            }
            catch (ArgumentException ex)
            {
                stderr.WriteLine(ex.Message);
                return 1;
            }
        }

        if (positional.Count > 1)
        {
            stderr.WriteLine("错误：load_skill 只接受一个 skill 名称。");
            return 1;
        }

        stderr.WriteLine("错误：使用 --path 时必须提供 skill 名称。");
        return 1;
    }
}
