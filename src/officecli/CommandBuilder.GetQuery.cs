// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildGetCommand(Option<bool> jsonOption)
    {
        var getFileArg = new Argument<FileInfo>("file") { Description = "Office 文档路径（即使使用 open/close mode 也必填）" };
        var pathArg = new Argument<string>("path") { Description = "DOM 路径（例如 /body/p[1]），或使用 'selected' 读取当前 watch 选择" };
        pathArg.DefaultValueFactory = _ => "/";
        var depthOpt = new Option<int>("--depth") { Description = "要包含的子节点深度" };
        depthOpt.DefaultValueFactory = _ => 1;
        var saveOpt = new Option<string?>("--save") { Description = "将底层二进制 payload（picture/ole/media）提取到此文件路径" };

        var getCommand = new Command("get", "按路径获取文档节点");
        getCommand.Add(getFileArg);
        getCommand.Add(pathArg);
        getCommand.Add(depthOpt);
        getCommand.Add(saveOpt);
        getCommand.Add(jsonOption);

        getCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(getFileArg)!;
            var path = MsysPathHint.Restore(result.GetValue(pathArg)!)!;
            var depth = result.GetValue(depthOpt);
            // CONSISTENCY(dos-hardening): cap user-supplied depth so a huge
            // --depth on a deeply-nested doc can't drive the node-building
            // recursion (and its O(n^2) InnerText/OuterXml-per-node cost) into
            // a multi-minute hang or stack overflow. See DocumentLimits.
            if (depth > DocumentLimits.MaxRecursionDepth)
                depth = DocumentLimits.MaxRecursionDepth;
            var savePath = result.GetValue(saveOpt);

            // Special pseudo-path "selected" — query the running watch process
            // for the currently-selected element paths and resolve them to nodes.
            if (string.Equals(path, "selected", StringComparison.OrdinalIgnoreCase))
            {
                return GetSelectedAction(file.FullName, depth, json);
            }

            if (TryResident(file.FullName, req =>
            {
                req.Command = "get";
                req.Json = json;
                req.Args["path"] = path;
                req.Args["depth"] = depth.ToString();
                if (!string.IsNullOrEmpty(savePath)) req.Args["save"] = savePath;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var node = handler.Get(path, depth);

            // CONSISTENCY(get-not-found-exit): some handler Get paths surface
            // "not found" via DocumentNode { Type = "error" } instead of
            // throwing (e.g. /numbering/abstractNum[@id=999]). Other paths
            // throw and exit 1 via SafeRun. Treat error-type nodes the same
            // way so callers get a consistent non-zero exit on missing paths.
            if (string.Equals(node.Type, "error", StringComparison.Ordinal))
            {
                var err = node.Text ?? $"Path not found: {path}";
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else
                    Console.Error.WriteLine($"Error: {err}");
                return 1;
            }

            // --save <path>: extract the binary payload backing an OLE /
            // picture / media node to disk. The handler exposes this via
            // TryExtractBinary which looks up the node's relId and copies
            // the part's stream. When the node has no backing binary, we
            // surface a clear error instead of silently succeeding.
            if (!string.IsNullOrEmpty(savePath))
            {
                if (!handler.TryExtractBinary(path, savePath, out var contentType, out var byteCount))
                {
                    var err = $"Node at '{path}' has no binary payload to extract (only ole/picture/media/embedded nodes can be saved).";
                    if (json)
                        Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                    else
                        Console.Error.WriteLine($"Error: {err}");
                    return 1;
                }
                node.Format["savedTo"] = savePath;
                node.Format["savedBytes"] = byteCount;
                if (!string.IsNullOrEmpty(contentType))
                    node.Format["savedContentType"] = contentType!;
            }

            if (json)
                // Unified envelope contract: single-path get returns the same
                // {matches, results: [...]} shape as `get selected` and `query`,
                // so agents and scripts can use one jq path everywhere. Text
                // mode keeps the rich single-node rendering.
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNodes(new List<DocumentNode> { node }, OutputFormat.Json)));
            else
                Console.WriteLine(OutputFormatter.FormatNode(node, OutputFormat.Text));
            return 0;
        }, json); });

        return getCommand;
    }

    private static int GetSelectedAction(string filePath, int depth, bool json)
    {
        var paths = WatchNotifier.QuerySelection(filePath);
        if (paths == null)
        {
            var msg = $"no watch running for {Path.GetFileName(filePath)}. Start one with: officecli watch \"{filePath}\"";
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg));
            else
                Console.Error.WriteLine($"Error: {msg}");
            return 1;
        }

        // Resolve each path to a DocumentNode. Skip paths that no longer exist
        // (e.g. element removed since selection was made) — silently drop them.
        var nodes = new List<OfficeCli.Core.DocumentNode>();
        if (paths.Length > 0)
        {
            using var handler = DocumentHandlerFactory.Open(filePath);
            foreach (var p in paths)
            {
                try
                {
                    var n = handler.Get(p, depth);
                    if (n != null) nodes.Add(n);
                }
                catch
                {
                    // path no longer resolves — drop
                }
            }
        }

        // Flatten row/column nodes into their children so text output is
        // grep-friendly (one cell per line instead of a single "/Sheet1/col[C]" line).
        var flat = new List<OfficeCli.Core.DocumentNode>();
        foreach (var n in nodes)
        {
            if (n.Children.Count > 0 && n.Type is "column" or "row")
                flat.AddRange(n.Children);
            else
                flat.Add(n);
        }

        if (json)
        {
            Console.WriteLine(OutputFormatter.WrapEnvelope(
                OutputFormatter.FormatNodes(flat, OutputFormat.Json)));
        }
        else
        {
            Console.WriteLine(OutputFormatter.FormatNodes(flat, OutputFormat.Text));
        }
        return 0;
    }

    private static Command BuildQueryCommand(Option<bool> jsonOption)
    {
        var queryFileArg = new Argument<FileInfo>("file") { Description = "Office 文档路径（即使使用 open/close mode 也必填）" };
        var selectorArg = new Argument<string>("selector") { Description = "CSS 风格 selector（例如 paragraph[style=Normal] > run[font!=Arial]）" };

        var queryFindOpt = new Option<string?>("--find") { Description = "将结果过滤为包含此文本的元素（不区分大小写的子串）" };
        var queryCompactOpt = new Option<bool>("--compact") { Description = "按文档顺序每个元素输出一行：path<TAB>[label]<TAB>\"text\"；文本最多 60 字（… 表示截断），空文本为 (empty)，table 折叠为 [table RxC]。末行固定为 total：PPTX 为 'total: N of M elements / K slides'，docx 为 'total: N of M elements'（不会出现容器后缀）；N 是上方元素行数（lineCount-1 == N 表示已完整读取），M 是全部顶级元素。完整列出文档时：PPTX 使用 selector '*'，docx 使用 'paragraph, table'，可使 N == M。label 是稳定集合：PPTX 为 title/placeholder/textbox/shape/picture/chart/connector/group/equation 加 'table RxC'；docx 为 style 名。该格式是稳定契约：已有列和 label 不会变更或重排，只会在末尾追加列。仅支持 pptx/docx；xlsx 请使用 'view text --range'。可用 --fields 追加 Format 字段。" };
        var queryFieldsOpt = new Option<string?>("--fields") { Description = "配合 --compact 在末尾追加的 Format 字段，逗号分隔（例如 x,y,width）；每项输出为 k=v，缺失值输出 k=" };

        var queryCommand = new Command("query", "使用 CSS 风格 selector 查询文档元素");
        queryCommand.Add(queryFileArg);
        queryCommand.Add(selectorArg);
        queryCommand.Add(jsonOption);
        queryCommand.Add(queryFindOpt);
        queryCommand.Add(queryCompactOpt);
        queryCommand.Add(queryFieldsOpt);

        queryCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(queryFileArg)!;
            var selector = MsysPathHint.Restore(result.GetValue(selectorArg)!)!;
            var textFilter = result.GetValue(queryFindOpt);
            var compact = result.GetValue(queryCompactOpt);
            var fields = result.GetValue(queryFieldsOpt);
            if (compact && json)
                throw new OfficeCli.Core.CliException("--compact 是纯文本逐行格式，不能与 --json 同时使用。") { Code = "invalid_value" };

            if (TryResident(file.FullName, req =>
            {
                req.Command = "query";
                req.Json = json;
                req.Args["selector"] = selector;
                if (textFilter != null) req.Args["find"] = textFilter;
                if (compact) req.Args["compact"] = "true";
                if (fields != null) req.Args["fields"] = fields;
            }, json) is {} rc) return rc;

            var format = json ? OutputFormat.Json : OutputFormat.Text;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            // CONSISTENCY(cell-selector-alias): the Excel cell selector accepts short
            // aliases (bold -> font.bold, size -> font.size, ...). FilterSelector
            // applies the same normalization, runs the boolean and/or engine, and
            // routes a pure-AND (flat) selector through the exact legacy path.
            // SelectorTargetsCells strips an optional sheet prefix (Sheet1!cell...,
            // /Sheet1/cell...) before the element check — without it, sheet-scoped
            // cell selectors skip alias normalization and drop all matches.
            Func<string, string>? keyResolver =
                handler is OfficeCli.Handlers.ExcelHandler
                && OfficeCli.Handlers.ExcelHandler.SelectorTargetsCells(selector)
                    ? OfficeCli.Handlers.ExcelHandler.ResolveCellAttributeAlias : null;
            var (results, warnings) = OfficeCli.Core.AttributeFilter.FilterSelector(selector, handler.Query, keyResolver);
            if (!string.IsNullOrEmpty(textFilter))
                results = results.Where(n => n.Text != null && OfficeCli.Core.AttributeFilter.MatchesTextFilter(n.Text, textFilter)).ToList();
            if (compact)
            {
                foreach (var w in warnings) Console.Error.WriteLine(w.Message);
                Console.WriteLine(FormatNodesCompact(handler, results, fields));
                return 0;
            }
            if (json)
            {
                // CONSISTENCY(query-json-children): Query returns nodes with empty
                // Children but populated ChildCount (handlers build query nodes at
                // depth=0 to avoid expensive subtree walks). For --json output we
                // hydrate children via Get(path, depth=1) so consumers see the same
                // shape that `get --json` produces.
                for (int i = 0; i < results.Count; i++)
                {
                    var n = results[i];
                    if (n.ChildCount > 0 && n.Children.Count == 0 && !string.IsNullOrEmpty(n.Path))
                    {
                        try
                        {
                            var hydrated = handler.Get(n.Path, depth: 1);
                            if (hydrated?.Children != null && hydrated.Children.Count > 0)
                                n.Children.AddRange(hydrated.Children);
                        }
                        catch { /* path may not be Get-resolvable; leave as-is */ }
                    }
                }
                var cliWarnings = warnings.Select(w => new OfficeCli.Core.CliWarning
                {
                    Message = w.Message,
                    Code = w.Code,
                    Kind = w.Kind,
                    Key = w.Key,
                    Value = w.Value,
                    Available = w.Available,
                    Suggestion = w.Suggestion,
                }).ToList();
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNodes(results, OutputFormat.Json),
                    cliWarnings.Count > 0 ? cliWarnings : null));
            }
            else
            {
                foreach (var w in warnings) Console.Error.WriteLine(w.Message);
                var output = OutputFormatter.FormatNodes(results, OutputFormat.Text);
                if (!string.IsNullOrEmpty(output))
                    Console.WriteLine(output);
                if (results.Count == 0)
                {
                    var ext = file.Extension.ToLowerInvariant().TrimStart('.');
                    Console.Error.WriteLine($"No matches. Run 'officecli {ext} query' for selector syntax.");
                }
            }
            return 0;
        }, json); });

        return queryCommand;
    }

    /// <summary>
    /// `query --compact` 的逐行稳定格式。此输出会被程序逐行解析并用于确认
    /// 读取完整性；列顺序、TAB 分隔符、`…` 截断标记、`(empty)`、`[label]`
    /// 形式和 total 行形状都是 API。已有列不得变更或重排，只允许末尾追加新列。
    ///
    ///   {path}\t[{label}]\t"{text ≤60 chars, \t/\n/\"/\\ escaped}"
    ///   {path}\t[table {R}x{C}]                     （table 折叠，无文本列）
    ///   {path}\t[{label}]\t(empty)                  （空文本）
    ///   --fields k1,k2 追加 \tk1=v1\tk2=v2          （缺失 key 输出 k=）
    ///   total: {N} of {M} elements / {K} slides     （pptx）
    ///   total: {N} of {M} elements                  （docx）
    ///
    /// N 恰好等于 total 行之前的元素行数（folded table 算 1），因此
    /// `lineCount - 1 == N` 表示读取了全部结果。M 是不受 selector 影响的
    /// 全部顶级元素（pptx 为全部 slide 中的 frame；docx 为 body-level block）。
    /// total 行始终存在（包括 N=0）且只在末尾出现一次。docx 的 total 行不带
    /// 容器后缀；该缺失本身也是冻结的格式契约，后续不能追加。
    ///
    /// 元素行遵循 document order：pptx 按 slide index 再按 z-order 排序，避免
    /// `*` 等多 type selector 按类型分组；docx 遵循 document flow。label 在
    /// 每个 format 中都处于固定位置：pptx 为 type 与折叠的 `table RxC`，docx
    /// 为 paragraph style。可以新增 label value，但已有 value 不会改变含义。
    /// </summary>
    internal static string FormatNodesCompact(IDocumentHandler handler, List<DocumentNode> results, string? fields)
    {
        if (handler is ExcelHandler)
            throw new OfficeCli.Core.CliException(
                "--compact 不支持 xlsx：请使用 view text，或使用 view text --range Sheet1!A1:C10。")
            { Code = "invalid_value" };

        var fieldList = string.IsNullOrWhiteSpace(fields)
            ? null
            : fields.Split(',').Select(f => f.Trim()).Where(f => f.Length > 0).ToList();

        if (handler is PowerPointHandler)
        {
            results = results
                .Select((n, i) => (n, i))
                .OrderBy(t => System.Text.RegularExpressions.Regex.Match(t.n.Path, @"^/slide\[(\d+)\]") is { Success: true } m
                    ? int.Parse(m.Groups[1].Value) : int.MaxValue)
                .ThenBy(t => t.n.Format.TryGetValue("zorder", out var z) && int.TryParse(z?.ToString(), out var zi)
                    ? zi : int.MaxValue)
                .ThenBy(t => t.i)
                .Select(t => t.n)
                .ToList();
        }

        var sb = new System.Text.StringBuilder();
        foreach (var n in results)
        {
            sb.Append(n.Path);
            if (n.Type == "table" && n.Format.TryGetValue("rows", out var r) && n.Format.TryGetValue("cols", out var c))
            {
                sb.Append('\t').Append($"[table {r}x{c}]");
            }
            else
            {
                var label = !string.IsNullOrEmpty(n.Style) ? n.Style : n.Type;
                sb.Append('\t').Append('[').Append(label).Append(']');
                sb.Append('\t').Append(CompactText(n.Text));
            }
            if (fieldList != null)
                foreach (var f in fieldList)
                    sb.Append('\t').Append(f).Append('=')
                      .Append(n.Format.TryGetValue(f, out var v) && v != null ? v.ToString() : "");
            sb.Append('\n');
        }

        var (total, containerSuffix) = CountCompactDenominator(handler);
        sb.Append($"total: {results.Count} of {total} elements{containerSuffix}");
        return sb.ToString();
    }

    private static string CompactText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return "(empty)";
        var t = text.Replace("\\", "\\\\").Replace("\t", "\\t")
                    .Replace("\r", "").Replace("\n", "\\n").Replace("\"", "\\\"");
        if (t.Length > 60) t = t[..60] + "…";
        return "\"" + t + "\"";
    }

    private static (int Total, string ContainerSuffix) CountCompactDenominator(IDocumentHandler handler)
    {
        if (handler is PowerPointHandler)
        {
            int slides = handler.Query("slide").Count;
            var paths = new HashSet<string>();
            foreach (var sel in new[] { "shape", "picture", "table", "chart", "connector", "group" })
            {
                try { foreach (var n in handler.Query(sel)) paths.Add(n.Path); }
                catch { /* selector unsupported on this document — skip */ }
            }
            return (paths.Count, $" / {slides} slides");
        }

        try
        {
            var body = handler.Get("/body", depth: 1);
            return (body.Children.Count(c => c.Type != "section"), "");
        }
        catch
        {
            return (0, "");
        }
    }
}
