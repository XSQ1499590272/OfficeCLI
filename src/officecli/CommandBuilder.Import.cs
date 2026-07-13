// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildImportCommand(Option<bool> jsonOption)
    {
        var importFileArg = new Argument<FileInfo>("file") { Description = "目标 Excel 文件（.xlsx）" };
        var importParentPathArg = new Argument<string>("parent-path") { Description = "Sheet 路径（例如 /Sheet1）" };
        var importSourceArg = new Argument<FileInfo?>("source-file") { Description = "要导入的源 CSV/TSV 文件（位置参数，可替代 --file）" };
        importSourceArg.DefaultValueFactory = _ => null!;
        var importSourceOpt = new Option<FileInfo?>("--file") { Description = "要导入的源 CSV/TSV 文件" };
        var importStdinOpt = new Option<bool>("--stdin") { Description = "从 stdin 读取 CSV/TSV 数据" };
        var importFormatOpt = new Option<string?>("--format") { Description = "数据格式：csv 或 tsv（默认从文件扩展名推断，否则为 csv）" };
        var importHeaderOpt = new Option<bool>("--header") { Description = "首行作为 header：设置 AutoFilter 并冻结窗格" };
        var importStartCellOpt = new Option<string>("--start-cell") { Description = "起始单元格（默认：A1）" };
        importStartCellOpt.DefaultValueFactory = _ => "A1";

        var importCommand = new Command("import", "将 CSV/TSV 数据导入 Excel sheet");
        importCommand.Add(importFileArg);
        importCommand.Add(importParentPathArg);
        importCommand.Add(importSourceArg);
        importCommand.Add(importSourceOpt);
        importCommand.Add(importStdinOpt);
        importCommand.Add(importFormatOpt);
        importCommand.Add(importHeaderOpt);
        importCommand.Add(importStartCellOpt);
        importCommand.Add(jsonOption);

        importCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(importFileArg)!;
            var parentPath = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(importParentPathArg)!)!;
            var source = result.GetValue(importSourceOpt) ?? result.GetValue(importSourceArg);
            var useStdin = result.GetValue(importStdinOpt);
            var format = result.GetValue(importFormatOpt);
            var header = result.GetValue(importHeaderOpt);
            var startCell = result.GetValue(importStartCellOpt)!;

            if (!file.Exists)
                throw new CliException($"File not found: {file.FullName}")
                {
                    Code = "file_not_found",
                    Suggestion = $"Create the file first: officecli create \"{file.FullName}\""
                };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".xlsx")
                throw new CliException("Import is only supported for .xlsx files in V1")
                {
                    Code = "unsupported_type",
                    Suggestion = "Use a .xlsx file"
                };

            // Read CSV content
            string csvContent;
            if (useStdin)
            {
                csvContent = Console.In.ReadToEnd();
            }
            else if (source != null)
            {
                if (!source.Exists)
                    throw new CliException($"Source file not found: {source.FullName}")
                    {
                        Code = "file_not_found"
                    };
                csvContent = File.ReadAllText(source.FullName, Encoding.UTF8);
            }
            else
            {
                throw new CliException("Either --file or --stdin must be specified")
                {
                    Code = "missing_argument",
                    Suggestion = "Use --file <path> to specify a CSV/TSV file, or --stdin to read from standard input"
                };
            }

            // Determine delimiter: --format flag > file extension > default csv
            char delimiter = ',';
            if (!string.IsNullOrEmpty(format))
            {
                delimiter = format.ToLowerInvariant() switch
                {
                    "tsv" => '\t',
                    "csv" => ',',
                    _ => throw new CliException($"Unknown format: {format}. Use 'csv' or 'tsv'")
                    {
                        Code = "invalid_value",
                        ValidValues = ["csv", "tsv"]
                    }
                };
            }
            else if (source != null)
            {
                var sourceExt = Path.GetExtension(source.FullName).ToLowerInvariant();
                if (sourceExt == ".tsv" || sourceExt == ".tab")
                    delimiter = '\t';
            }

            // Release any running resident's file lock before direct-open (import bypasses resident)
            ResidentClient.SendClose(file.FullName);
            using var handler = new OfficeCli.Handlers.ExcelHandler(file.FullName, editable: true);
            var msg = handler.Import(parentPath, csvContent, delimiter, header, startCell);
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
            else
                Console.WriteLine(msg);
            return 0;
        }, json); });

        return importCommand;
    }

    private static Command BuildCreateCommand(Option<bool> jsonOption)
    {
        var createFileArg = new Argument<string>("file") { Description = "输出文件路径（.docx、.xlsx、.pptx）" };
        var createTypeOpt = new Option<string>("--type") { Description = "文档类型（docx、xlsx、pptx）；可选，默认从文件扩展名推断" };
        var createForceOpt = new Option<bool>("--force") { Description = "覆盖已存在的文件。" };
        var createLocaleOpt = new Option<string>("--locale") { Description = "Locale tag（例如 zh-CN、ja、ko、ar、he）：在 docDefaults 中设置各文字系统默认字体，并为 Arabic / Hebrew / Persian / Urdu 等 RTL locale 启用 RTL 布局。不传时默认使用 OS 用户 culture（macOS 的 CFLocale、Linux 的 $LANG、Windows 的用户 UI culture）。传入 --locale en-US 可不受宿主机影响，强制确定性的 LTR/Latin 基线。目前仅对 .docx 生效。" };
        var createMinimalOpt = new Option<bool>("--minimal") { Description = "（仅 .docx）跳过 Word 的 Normal.dotm 风格基线（Calibri 11pt + Normal style + theme1.xml），改为输出原始 OOXML-spec docx。适用于测试边界场景或生成最精简输出。不传时文档采用与 Word 对齐的默认值，以便在 Word、其他 producer 和 CLI preview 中一致渲染。" };
        var createCommand = new Command("create", "创建空白 Office 文档");
        createCommand.Add(createFileArg);
        createCommand.Add(createTypeOpt);
        createCommand.Add(createForceOpt);
        createCommand.Add(createLocaleOpt);
        createCommand.Add(createMinimalOpt);
        createCommand.Add(jsonOption);

        createCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(createFileArg)!;
            var type = result.GetValue(createTypeOpt);
            var force = result.GetValue(createForceOpt);
            var explicitLocale = result.GetValue(createLocaleOpt);
            var minimal = result.GetValue(createMinimalOpt);

            // Fall back to OS user culture when --locale is not explicitly
            // given. Empty / C / POSIX cultures yield null (no locale baked)
            // so CI environments produce neutral output.
            var locale = OfficeCli.Core.LocaleFontRegistry.ResolveEffectiveLocale(explicitLocale);
            bool localeInferred = !string.IsNullOrWhiteSpace(locale)
                && string.IsNullOrWhiteSpace(explicitLocale);

            // If file has no extension but --type is provided, append it
            if (!string.IsNullOrEmpty(type) && string.IsNullOrEmpty(Path.GetExtension(file)))
            {
                var ext = type.StartsWith('.') ? type : "." + type;
                file += ext;
            }

            // Check if the file is held by a resident process
            var fullPath = Path.GetFullPath(file);
            if (ResidentClient.TryConnect(fullPath, out _))
            {
                // Stale-resident recovery: if the on-disk file is gone (typical
                // example-script pattern: os.remove(FILE) then `create`), the
                // resident is pinning a path that no longer exists. Auto-close
                // it and proceed — refusing here would force every example
                // script to wrap `create` in a defensive `close`.
                if (!File.Exists(fullPath))
                {
                    ResidentClient.SendClose(fullPath);
                }
                else
                {
                    throw new CliException($"{Path.GetFileName(file)} is currently opened by a resident process. Please run 'officecli close \"{file}\"' first.")
                    {
                        Code = "file_locked",
                        Suggestion = $"Run: officecli close \"{file}\""
                    };
                }
            }

            // Refuse to silently overwrite an existing file unless --force is set.
            // OpenXML SDK's Create truncates the target otherwise, which can destroy
            // user data when an AI agent retries or mis-types the path.
            if (File.Exists(fullPath) && !force)
            {
                throw new CliException($"File already exists: {file}. Use --force to overwrite.")
                {
                    Code = "file_exists",
                    Suggestion = "Add --force flag or remove the file first."
                };
            }
            if (File.Exists(fullPath) && force)
            {
                Console.Error.WriteLine($"Overwriting existing file: {file}");
            }

            OfficeCli.BlankDocCreator.Create(file, locale, minimal);
            var fullCreatedPath = Path.GetFullPath(file);

            // If a --force overwrite replaced a file that currently has a live
            // watch session, push a full SSE refresh so the preview reflects the
            // new (blank) document instead of the stale pre-overwrite content
            // (issue #169). create replaces the whole file, so a full re-render
            // is the only correct shape — mirrors swap / refresh. Only reachable
            // when no resident pins the file (otherwise create fails file_locked
            // above); the watch server itself never opens the file. Best-effort:
            // a preview-refresh failure must never fail the create itself.
            if (WatchServer.IsWatching(fullCreatedPath))
            {
                try
                {
                    using var watchHandler = OfficeCli.Handlers.DocumentHandlerFactory.Open(fullCreatedPath, editable: false);
                    NotifyWatch(watchHandler, fullCreatedPath, null);
                }
                catch { /* preview refresh is best-effort; the file is already written */ }
            }

            // Best-effort: auto-start a short-lived resident process so
            // follow-up commands on this freshly-created file hit the
            // in-memory handler instead of re-opening from disk each time.
            // Uses a 60s idle timeout (much shorter than `open`'s default
            // 12min) so a stray `create` with no follow-up exits quickly.
            // Failure here does NOT fail the command — the file is already
            // on disk and all other commands still work via direct open.
            var noAuto = Environment.GetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT");
            string? residentErr = null;
            var residentStarted = noAuto == "1" || string.Equals(noAuto, "true", StringComparison.OrdinalIgnoreCase)
                ? false
                : TryStartResidentProcess(fullCreatedPath, idleSeconds: 60, out residentErr);
            var residentSuffix = residentStarted
                ? " (kept open in background for faster subsequent commands)"
                : "";

            if (json)
            {
                Console.WriteLine(OutputFormatter.WrapEnvelopeText($"Created: {fullCreatedPath}{residentSuffix}"));
            }
            else
            {
                Console.WriteLine($"Created: {file}{residentSuffix}");
                // Surface the inferred locale on stderr so the user can see
                // when the OS culture shaped the doc (RTL layout, CJK fonts,
                // etc.). Stays out of stdout / JSON envelope so scripts that
                // pipe `create` output aren't disturbed.
                if (localeInferred && Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    var rtlNote = OfficeCli.Core.LocaleFontRegistry.IsRightToLeft(locale) ? " (RTL layout enabled)" : "";
                    Console.Error.WriteLine($"Note: locale '{locale}' inferred from OS user culture{rtlNote}. Pass --locale to override.");
                }
                if (!residentStarted && !string.IsNullOrEmpty(residentErr))
                {
                    Console.Error.WriteLine($"Note: resident auto-start failed ({residentErr}); falling back to direct file access.");
                }
                if (Path.GetExtension(file).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"  totalSlides: 0");
                    // Pair the unit so both dimensions agree (matches Get /
                    // readback after R40 — paired emit avoids mixing pt+cm).
                    var (cWStr, cHStr) = Core.EmuConverter.FormatEmuPaired(12192000, 6858000);
                    Console.WriteLine($"  slideWidth: {cWStr}");
                    Console.WriteLine($"  slideHeight: {cHStr}");
                }
            }
            return 0;
        }, json); });

        return createCommand;
    }

    private static Command BuildMergeCommand(Option<bool> jsonOption)
    {
        var mergeTemplateArg = new Argument<string>("template") { Description = "含 {{key}} 占位符的 template 文件路径（.docx、.xlsx、.pptx）" };
        var mergeOutputArg = new Argument<string>("output") { Description = "输出文件路径" };
        var mergeDataOpt = new Option<string>("--data") { Description = "JSON 数据或 .json 文件路径", Required = true };
        var mergeForceOpt = new Option<bool>("--force") { Description = "覆盖已存在的输出文件。" };
        var mergeCommand = new Command("merge", "以 JSON 数据合并 template，替换 {{key}} 占位符");
        mergeCommand.Add(mergeTemplateArg);
        mergeCommand.Add(mergeOutputArg);
        mergeCommand.Add(mergeDataOpt);
        mergeCommand.Add(mergeForceOpt);
        mergeCommand.Add(jsonOption);

        mergeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var template = result.GetValue(mergeTemplateArg)!;
            var output = result.GetValue(mergeOutputArg)!;
            var dataArg = result.GetValue(mergeDataOpt)!;
            var force = result.GetValue(mergeForceOpt);

            // If a resident holds the template with unsaved in-memory edits,
            // ask it to flush to disk first — otherwise File.Copy reads the
            // stale pre-edit bytes (resident saves itself; we never open the
            // file here). Mirrors import's pre-direct-open resident handshake.
            ResidentClient.SendSave(Path.GetFullPath(template));

            var data = Core.TemplateMerger.ParseMergeData(dataArg);
            var mergeResult = Core.TemplateMerger.Merge(template, output, data, force);

            if (json)
            {
                var mergeData = new System.Text.Json.Nodes.JsonObject
                {
                    ["output"] = Path.GetFullPath(output),
                    ["replacedKeys"] = mergeResult.UsedKeys.Count,
                    ["unresolvedPlaceholders"] = new System.Text.Json.Nodes.JsonArray(
                        mergeResult.UnresolvedPlaceholders.Select(p => (System.Text.Json.Nodes.JsonNode)p).ToArray())
                };
                var unresolvedCount = mergeResult.UnresolvedPlaceholders.Count;
                var message = unresolvedCount > 0
                    ? $"Merged {mergeResult.UsedKeys.Count} key(s), {unresolvedCount} unresolved placeholder(s)"
                    : $"Merged {mergeResult.UsedKeys.Count} key(s)";
                var jsonObj = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = mergeData,
                    ["message"] = message,
                };
                Console.WriteLine(jsonObj.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = false }));
            }
            else
            {
                Console.WriteLine($"Merged: {output}");
                Console.WriteLine($"  Replaced keys: {mergeResult.UsedKeys.Count}");
                if (mergeResult.UnresolvedPlaceholders.Count > 0)
                {
                    Console.Error.WriteLine($"  Warning: {mergeResult.UnresolvedPlaceholders.Count} unresolved placeholder(s):");
                    foreach (var p in mergeResult.UnresolvedPlaceholders)
                        Console.Error.WriteLine($"    - {{{{{p}}}}}");
                }
            }
            return 0;
        }, json); });

        return mergeCommand;
    }
}
