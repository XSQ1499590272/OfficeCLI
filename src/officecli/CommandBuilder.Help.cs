// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Help;

namespace OfficeCli;

static partial class CommandBuilder
{
    // Recognized verbs that route help through the operation-scoped filter.
    // Matches IDocumentHandler's public surface — keep in sync if new verbs
    // are added to the handler API.
    private static readonly string[] HelpVerbs =
        { "add", "set", "get", "query", "remove" };

    // MCP and load_skill are dispatched before System.CommandLine sees them,
    // so help renders their usage directly.
    /// <summary>
    /// Print the verbose usage block for an early-dispatch command
    /// (mcp/load_skill) to the given writer. Single source of truth shared
    /// between `officecli help &lt;cmd&gt;`, the integration stubs' SetAction, and
    /// Program.cs's invalid-args error path. Returns true if the command name
    /// was recognized.
    /// </summary>
    internal static bool WriteEarlyDispatchUsage(string name, TextWriter writer)
    {
        if (!EarlyDispatchHelp.TryGetValue(name, out var lines)) return false;
        foreach (var line in lines) writer.WriteLine(line);
        return true;
    }

    private static readonly Dictionary<string, string[]> EarlyDispatchHelp =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["mcp"] = new[]
            {
                "用法：",
                "  officecli mcp                    启动 MCP stdio server（供 AI agent 使用）",
                "  officecli mcp <target>           向 MCP client 注册 officecli",
                "  officecli mcp uninstall <target> 从 MCP client 注销 officecli",
                "  officecli mcp list               查看各 client 的注册状态",
                "",
                "目标：lms（LM Studio）、claude（Claude Code）、cursor、vscode（Copilot）",
            },
            ["load_skill"] = new[]
            {
                "用法：",
                "  officecli load_skill                         列出全部 Skill 及其触发条件",
                "  officecli load_skill <name>                 输出该 Skill 的 SKILL.md 及内嵌参考文件清单",
                "  officecli load_skill <name> --path <relpath> 输出一个内嵌参考文件（例如 --path reference/decision-rules.md）",
                "",
                "Skill：pptx、word、excel、word-form、morph-ppt、morph-ppt-3d、pitch-deck、academic-paper、data-dashboard、financial-model",
                "二进制参考资源不能通过文本通道输出。",
            },
        };

    /// <summary>
    /// `officecli help [format] [verb] [element] [--json]` — schema-driven help.
    ///
    /// Argument forms accepted:
    ///   help                         → list formats
    ///   help &lt;format&gt;                → list all elements
    ///   help &lt;format&gt; &lt;verb&gt;         → list elements supporting that verb
    ///   help &lt;format&gt; &lt;element&gt;      → full element detail
    ///   help &lt;format&gt; &lt;verb&gt; &lt;element&gt; → verb-filtered element detail
    ///
    /// The middle arg is interpreted as verb iff it matches HelpVerbs.
    /// Mirrors the actual CLI structure: `officecli &lt;verb&gt; &lt;file&gt; ...`, so
    /// `officecli help docx add chart` reads exactly like the command you
    /// are about to run.
    /// </summary>
    public static Command BuildHelpCommand(Option<bool> jsonOption, RootCommand? rootCommand = null)
    {
        var formatArg = new Argument<string?>("format")
        {
            Description = "文档格式：docx/xlsx/pptx（别名：word、excel、ppt、powerpoint）。省略时列出格式。",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var secondArg = new Argument<string?>("verb-or-element")
        {
            Description = "动词（add/set/get/query/remove）或元素名称。省略时列出全部元素。",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var thirdArg = new Argument<string?>("element")
        {
            Description = "已给出动词时的元素名称（例如 'help docx add chart'）。",
            Arity = ArgumentArity.ZeroOrOne,
        };
        // Scoped to `help` only — `help all`/`help <fmt> all` can emit either:
        //   --json   one envelope-wrapped JSON document (matches other CLI
        //            commands; one parse for the whole corpus)
        //   --jsonl  NDJSON (one self-contained JSON object per line, no
        //            envelope, streaming-friendly)
        // Mutually exclusive on `help all`. Other help forms ignore --jsonl
        // since they're either single documents (use --json) or human-readable
        // listings with no JSON form.
        var jsonlOption = new Option<bool>("--jsonl")
        {
            Description = "（仅 help all）输出 NDJSON：每行一个 JSON object，不加 envelope。",
        };

        var command = new Command("help", "显示 officecli 基于 schema 的命令参考。");
        command.Add(formatArg);
        command.Add(secondArg);
        command.Add(thirdArg);
        command.Add(jsonOption);
        command.Add(jsonlOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            var jsonl = result.GetValue(jsonlOption);
            var format = result.GetValue(formatArg);
            var second = result.GetValue(secondArg);
            var third = result.GetValue(thirdArg);

            // Disambiguate middle arg: is it a verb or an element?
            string? verb = null;
            string? element = null;
            if (second != null)
            {
                if (third != null)
                {
                    // 3 args: format, verb, element — second is a verb only if it
                    // actually looks like one. If format is itself a HelpVerb (from
                    // the `<cmd> --help <format> <element>` rewrite) then second is
                    // a document format token, not a verb; leave verb=null so Case 1b
                    // handles it by showing SCL help for the command.
                    // CONSISTENCY(args-rewrite): mirrors the 2-arg guard below.
                    if (HelpVerbs.Contains(second, StringComparer.OrdinalIgnoreCase))
                    {
                        verb = second;
                        element = third;
                    }
                    else if (SchemaHelpLoader.IsKnownFormat(format!))
                    {
                        // format is a real schema format AND third is provided, but
                        // second isn't a verb — surface the error instead of
                        // silently falling through to Case 2 (which would list all
                        // elements, ignoring user input).
                        Console.Error.WriteLine(
                            $"错误：未知动词“{second}”。可用值：{string.Join(", ", HelpVerbs)}。");
                        return 1;
                    }
                    // else: format is a HelpVerb (CRUD-verb-as-format from the
                    // `<verb> --help <fmt> <element>` rewrite), second is the format
                    // token, third is the element — fall through with verb=null,
                    // element=null so Case 1b shows SCL command help.
                }
                else if (HelpVerbs.Contains(second, StringComparer.OrdinalIgnoreCase))
                {
                    // 2 args where second is a verb: filter listing by verb.
                    verb = second;
                }
                else
                {
                    // 2 args where second is NOT a verb: treat as element.
                    element = second;
                }
            }

            return SafeRun(() => RunHelp(format, verb, element, json, jsonl, rootCommand), json);
        });

        return command;
    }

    private static int RunHelp(string? format, string? verb, string? element, bool json, bool jsonl, RootCommand? rootCommand)
    {
        // --json and --jsonl are mutually exclusive on `help all` / `help <fmt>
        // all`: the first emits one envelope-wrapped JSON document, the second
        // emits NDJSON. Combining them has no coherent meaning. Reject early
        // with a clear message rather than silently picking one.
        if (json && jsonl)
        {
            Console.Error.WriteLine("错误：--json 与 --jsonl 不能同时使用。");
            return 1;
        }

        // Case 1: no args — print SCL's default help (Description, Usage,
        // Options, full Commands list with arg signatures + descriptions),
        // then append the schema-driven reference block. The SCL output is
        // the single source of truth for the command surface; this command
        // only adds what SCL doesn't know about (formats, schema verbs,
        // aliases, drill-in usage).
        // Use `== null` (not IsNullOrEmpty) so an explicit empty-string format
        // (`help '' docx paragraph`) falls through to NormalizeFormat → proper
        // "unknown format ''" error, instead of silently discarding the
        // trailing tokens by routing into the no-args banner.
        // CONSISTENCY(empty-arg) — mirrors the Case 2 element guard.
        // Case 0: `help all` — flat, grep-friendly dump of every (format,
        // element, property) row across the schema corpus. One self-contained
        // line per record so `officecli help all | grep <term>` returns
        // intelligible matches without context loss.
        if (string.Equals(format, "all", StringComparison.OrdinalIgnoreCase))
        {
            if (verb != null || element != null)
            {
                Console.Error.WriteLine(
                    "错误：'help all' 不接受额外参数。可通过管道传给 grep 筛选。");
                return 1;
            }
            if (json)
            {
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    SchemaHelpFlatRenderer.RenderAllJsonArray()));
                return 0;
            }
            Console.Write(jsonl
                ? SchemaHelpFlatRenderer.RenderAllJsonl()
                : SchemaHelpFlatRenderer.RenderAll());
            return 0;
        }

        // Case 0b: `help <format> all` — same flat dump but filtered to one
        // format. "all" isn't a CRUD verb so it lands in `element` after the
        // upstream disambiguation. Saves the user a `| grep ^<format>`.
        if (format != null
            && SchemaHelpLoader.IsKnownFormat(format)
            && verb == null
            && string.Equals(element, "all", StringComparison.OrdinalIgnoreCase))
        {
            var canonical = SchemaHelpLoader.NormalizeFormat(format);
            if (json)
            {
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    SchemaHelpFlatRenderer.RenderAllJsonArray(canonical)));
                return 0;
            }
            Console.Write(jsonl
                ? SchemaHelpFlatRenderer.RenderAllJsonl(canonical)
                : SchemaHelpFlatRenderer.RenderAll(canonical));
            return 0;
        }

        if (format == null)
        {
            if (rootCommand != null)
            {
                // rootCommand.Parse(["--help"]) routes to SCL's HelpOption,
                // which writes Description/Usage/Options/Commands directly to
                // Console. Note Program.cs's `--help` → `help` rewrite only
                // runs once at process startup on the original args, so this
                // programmatic Parse goes straight to SCL and does not loop.
                rootCommand.Parse(new[] { "--help" }).Invoke();
                Console.WriteLine();
            }

            Console.WriteLine("Schema 参考（docx/xlsx/pptx）：");
            Console.WriteLine("  officecli help <format>                         列出全部元素");
            Console.WriteLine("  officecli help <format> <verb>                  列出支持该动词的元素");
            Console.WriteLine("  officecli help <format> <element>               显示完整元素详情");
            Console.WriteLine("  officecli help <format> <verb> <element>        显示该动词范围内的元素详情");
            Console.WriteLine("  officecli help <format> <element> --json        输出原始 schema JSON");
            Console.WriteLine("  officecli help all                              平铺列出全部（format、element、property），可通过管道传给 grep");
            Console.WriteLine("  officecli help all --json                       将相同内容输出为一个 envelope 包装的 JSON 文档");
            Console.WriteLine("  officecli help all --jsonl                      将相同内容输出为 NDJSON（每行一个 JSON object）");
            Console.WriteLine();
            Console.Write("  格式：");
            Console.WriteLine(string.Join(", ", SchemaHelpLoader.ListFormats()));
            Console.WriteLine("  动词：add、set、get、query、remove");
            Console.WriteLine("  别名：word→docx、excel→xlsx、ppt/powerpoint→pptx");
            Console.WriteLine();
            Console.WriteLine("提示：多数 shell 会展开 [brackets]，请为路径加引号：officecli get doc.docx \"/body/p[1]\"");
            return 0;
        }

        // Case 1b: not a format — try command help.
        //   - The early-dispatch MCP command does not participate in the
        //     command tree, so print this text before normal parsing.
        //     a hardcoded usage blurb.
        //   - Registered SCL subcommands get their --help forwarded.
        //
        // CONSISTENCY(args-rewrite): `officecli set --help chart` is rewritten to
        // `officecli help set chart` by Program.cs. "set" is not a document format,
        // so we fall into this branch. The trailing element token ("chart") has no
        // meaning in SCL command-help context — ignore it and show SCL help for "set".
        // Guard drops `element == null` for CRUD verbs so the rewrite case is handled.
        if (!SchemaHelpLoader.IsKnownFormat(format)
            && verb == null
            && (element == null || HelpVerbs.Contains(format, StringComparer.OrdinalIgnoreCase)
                || EarlyDispatchHelp.ContainsKey(format)))
        {
            if (WriteEarlyDispatchUsage(format, Console.Out))
                return 0;

            if (rootCommand != null)
            {
                var match = rootCommand.Subcommands.FirstOrDefault(
                    c => string.Equals(c.Name, format, StringComparison.OrdinalIgnoreCase)
                         && !c.Hidden
                         && c.Name != "help");
                if (match != null)
                    return rootCommand.Parse(new[] { match.Name, "--help" }).Invoke();
            }
        }

        // Validate verb if supplied.
        if (verb != null && !HelpVerbs.Contains(verb, StringComparer.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine($"错误：未知动词“{verb}”。可用值：{string.Join(", ", HelpVerbs)}。");
            return 1;
        }

        var canonicalFormat = SchemaHelpLoader.NormalizeFormat(format);

        // Case 2: format (+ optional verb) only — list elements.
        // Use `== null` (not IsNullOrEmpty) so that an explicit empty-string
        // arg (`help docx ''`) falls through to Case 3 where LoadSchema raises
        // a proper "unknown element ''" error. CONSISTENCY(empty-arg).
        if (element == null)
        {
            var all = SchemaHelpLoader.ListElements(canonicalFormat);
            var filtered = verb == null
                ? all
                : all.Where(el => SchemaHelpLoader.ElementSupportsVerb(canonicalFormat, el, verb!)).ToList();

            if (filtered.Count == 0 && verb != null)
            {
                Console.WriteLine($"{canonicalFormat} 中没有支持“{verb}”的元素。");
                return 0;
            }

            var header = verb == null
                ? $"{canonicalFormat} 的元素："
                : $"{canonicalFormat} 中支持“{verb}”的元素：";
            Console.WriteLine(header);

            // Build parent → children map for tree rendering. Children whose
            // declared parent isn't itself in the filtered set float back up
            // to top-level so nothing disappears under a filter.
            var filteredSet = new HashSet<string>(filtered, StringComparer.Ordinal);
            var parentOf = filtered.ToDictionary(
                el => el,
                el => SchemaHelpLoader.GetParentForTree(canonicalFormat, el),
                StringComparer.Ordinal);

            var topLevel = new List<string>();
            var byParent = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            foreach (var el in filtered)
            {
                var pr = parentOf[el];
                if (pr != null && filteredSet.Contains(pr))
                {
                    if (!byParent.TryGetValue(pr, out var list))
                        byParent[pr] = list = new List<string>();
                    list.Add(el);
                }
                else
                {
                    topLevel.Add(el);
                }
            }

            void WriteNode(string el, int depth)
            {
                Console.WriteLine($"{new string(' ', 2 + depth * 2)}{el}");
                if (byParent.TryGetValue(el, out var kids))
                    foreach (var kid in kids)
                        WriteNode(kid, depth + 1);
            }
            foreach (var el in topLevel)
                WriteNode(el, 0);
            Console.WriteLine();

            var detailHint = verb == null
                ? $"运行 'officecli help {canonicalFormat} <element>' 查看详情。"
                : $"运行 'officecli help {canonicalFormat} {verb} <element>' 查看按动词筛选的详情。";
            Console.WriteLine(detailHint);
            return 0;
        }

        // Case 3: format + (optional verb) + element — render schema.
        using var doc = SchemaHelpLoader.LoadSchema(format, element);
        Console.WriteLine(json
            ? SchemaHelpRenderer.RenderJson(doc)
            : SchemaHelpRenderer.RenderHuman(doc, verb));
        return 0;
    }

}
