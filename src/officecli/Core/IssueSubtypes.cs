// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Central catalogue of <c>view issues --type</c> accepted values. Single
/// source of truth so the CLI front-end (CommandBuilder.View) and the
/// resident server (ResidentServer.ExecuteView) reject typos identically
/// and the cross-handler protocol documentation cannot drift from what the
/// validator actually accepts.
/// </summary>
public static class IssueSubtypes
{
    public const string FormulaNotEvaluated = "formula_not_evaluated";
    public const string FormulaCacheStale = "formula_cache_stale";
    public const string FormulaRefMissingSheet = "formula_ref_missing_sheet";
    public const string FormulaEvalError = "formula_eval_error";
    public const string FieldNotEvaluated = "field_not_evaluated";
    public const string FieldCacheStale = "field_cache_stale";
    public const string SlideFieldNotEvaluated = "slide_field_not_evaluated";
    public const string ChartSeriesRefMissingSheet = "chart_series_ref_missing_sheet";
    public const string ChartCacheStale = "chart_cache_stale";
    public const string DefinedNameBroken = "definedname_broken";
    public const string DefinedNameTargetMissing = "definedname_target_missing";
    public const string BrokenPartRef = "broken_part_ref";
    /// <summary>pptx-only: notesSlide raw-set passthrough references an
    /// rId (<c>r:embed</c> / <c>r:link</c>) the dump pass cannot reproduce
    /// on the replay target (e.g. a non-image rel attached to a NotesSlidePart
    /// — embedded media, OLE, etc.). The raw-set still emits, but PowerPoint
    /// shows the referenced object as a broken placeholder on open. Emitted
    /// as an UnsupportedWarning during dump; the surfaced site is the slide
    /// owning the notes (<c>/slide[N]/notes</c>).</summary>
    public const string NotesUnresolvedRid = "notes_unresolved_rid";
    /// <summary>pptx-only: a shape with its own opaque dark solid fill carries
    /// opaque dark text (fill brightness &lt; 30%, run brightness &lt; 80%) — the
    /// text is unreadable when projected. Declared-model only: the shape's
    /// explicit fill is compared against its explicit run colors, so the
    /// backdrop is unambiguous (no z-order guesswork). Scheme/inherited colors,
    /// translucent runs, and colors carrying lumMod/shade transforms are
    /// skipped to keep false positives near zero. Format bucket, Warning.</summary>
    public const string LowContrast = "low_contrast";

    /// <summary>Broad IssueType bucket names — the canonical surface shown
    /// in error messages and help. Single-letter aliases (<see cref="BucketAliases"/>)
    /// are accepted by Validate but kept out of the user-facing list so the
    /// canonical-vs-alias distinction is visible.</summary>
    public static readonly string[] BucketNames =
        new[] { "format", "content", "structure" };

    /// <summary>Single-letter aliases accepted in addition to the canonical
    /// bucket names. Kept separate from <see cref="BucketNames"/> so error
    /// listings don't expose them as first-class values.</summary>
    public static readonly string[] BucketAliases =
        new[] { "f", "c", "s" };

    /// <summary>Combined accepted bucket inputs (canonical + aliases).</summary>
    public static readonly string[] ValidBuckets =
        BucketNames.Concat(BucketAliases).ToArray();

    /// <summary>Every subtype the <c>view issues</c> filter accepts by name.</summary>
    public static readonly string[] ValidSubtypes = new[]
    {
        FormulaNotEvaluated, FormulaCacheStale, FormulaRefMissingSheet, FormulaEvalError,
        FieldNotEvaluated, FieldCacheStale,
        SlideFieldNotEvaluated, NotesUnresolvedRid, LowContrast,
        ChartSeriesRefMissingSheet, ChartCacheStale,
        DefinedNameBroken, DefinedNameTargetMissing,
        BrokenPartRef,
    };

    /// <summary>Subtypes that are scanned by default and surface under
    /// <c>--type content</c>. Opt-in subtypes (currently only
    /// <see cref="ChartCacheStale"/>) require an exact-name request.</summary>
    public static readonly string[] OptInSubtypes = new[] { ChartCacheStale };

    /// <summary>One-line summary suitable for the CLI <c>--type</c> help
    /// text. Generated from <see cref="ValidSubtypes"/> so the help cannot
    /// drift from the validator.</summary>
    public static string TypeHelpDescription()
    {
        var defaults = ValidSubtypes.Where(s => !OptInSubtypes.Contains(s));
        return "问题类型过滤器。宽泛分类："
            + string.Join(", ", BucketNames)
            + "（别名：" + string.Join(", ", BucketAliases) + "）。"
            + "子类型（Content 分类；默认返回，也可通过 --type content 获取）："
            + string.Join(", ", defaults) + "。"
            + "仅显式启用（按准确名称请求；不包含在 --type content 中）："
            + string.Join(", ", OptInSubtypes) + "。"
            + "子类型与格式相关：formula_* / chart_* / definedname_* 适用于 xlsx，"
            + "field_* 适用于 docx，slide_field_* / notes_unresolved_rid / broken_part_ref / low_contrast 适用于 pptx；请求不适用于"
            + "当前文件的子类型会返回 count=0（不是错误）。"
            + "所有值均不区分大小写，并会去除首尾空白。";
    }

    /// <summary>
    /// Validate a user-supplied <c>--type</c> argument and return the
    /// canonicalised form. Null, empty, and whitespace-only inputs are
    /// normalised to null (treated as "no filter"). Surrounding whitespace
    /// is trimmed so values copied from shells with extra spaces still
    /// match. Recognised buckets and subtypes (case-insensitive) pass
    /// through unchanged. Anything else raises <see cref="CliException"/>
    /// with the full valid list — turning silent typos into a clear
    /// failure on both the CLI front-end and the resident-server fan-out.
    /// </summary>
    public static string? Validate(string? issueType)
    {
        if (string.IsNullOrWhiteSpace(issueType)) return null;
        var trimmed = issueType.Trim();
        var canonical = trimmed.ToLowerInvariant();
        foreach (var v in ValidBuckets) if (v == canonical) return trimmed;
        foreach (var v in ValidSubtypes) if (v == canonical) return trimmed;
        var all = ValidBuckets.Concat(ValidSubtypes).ToArray();
        throw new CliException(
            $"Invalid --type value: '{issueType}'. Valid buckets: {string.Join(", ", BucketNames)} (alias {string.Join(", ", BucketAliases)}). Valid subtypes: {string.Join(", ", ValidSubtypes)}.")
        { Code = "invalid_issue_type", ValidValues = all };
    }
}
