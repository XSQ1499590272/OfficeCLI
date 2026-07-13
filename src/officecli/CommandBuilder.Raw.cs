// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildRawCommand(Option<bool> jsonOption)
    {
        var rawFileArg = new Argument<FileInfo>("file") { Description = "Office 文档路径（即使使用 open/close mode 也必填）" };
        var rawPathArg = new Argument<string>("part") { Description = "Part 路径（例如 /document、/styles、/header[1]）" };
        rawPathArg.DefaultValueFactory = _ => "/document";

        var rawStartOpt = new Option<int?>("--start") { Description = "起始行号（仅 Excel sheet）" };
        var rawEndOpt = new Option<int?>("--end") { Description = "结束行号（仅 Excel sheet）" };

        var rawColsOpt = new Option<string?>("--cols") { Description = "列过滤器，使用逗号分隔（仅 Excel，例如 A,B,C）" };

        var rawCommand = new Command("raw", "查看文档 part 的原始 XML");
        rawCommand.Add(rawFileArg);
        rawCommand.Add(rawPathArg);
        rawCommand.Add(rawStartOpt);
        rawCommand.Add(rawEndOpt);
        rawCommand.Add(rawColsOpt);
        rawCommand.Add(jsonOption);

        rawCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(rawFileArg)!;
            var partPath = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(rawPathArg)!)!;
            var startRow = result.GetValue(rawStartOpt);
            var endRow = result.GetValue(rawEndOpt);
            var rawColsStr = result.GetValue(rawColsOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "raw";
                req.Args["part"] = partPath;
                if (startRow.HasValue) req.Args["start"] = startRow.Value.ToString();
                if (endRow.HasValue) req.Args["end"] = endRow.Value.ToString();
                if (rawColsStr != null) req.Args["cols"] = rawColsStr;
            }, json) is {} rc) return rc;

            var rawCols = rawColsStr != null ? new HashSet<string>(rawColsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var xml = handler.Raw(partPath, startRow, endRow, rawCols);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(xml));
            else Console.WriteLine(xml);
            return 0;
        }, json); });

        return rawCommand;
    }

    private static Command BuildRawSetCommand(Option<bool> jsonOption)
    {
        var rawSetFileArg = new Argument<FileInfo>("file") { Description = "Office 文档路径（即使使用 open/close mode 也必填）" };
        var rawSetPartArg = new Argument<string>("part") { Description = "Part 路径（例如 /document、/styles、/Sheet1、/slide[1]）" };
        var rawSetXpathOpt = new Option<string>("--xpath") { Description = "目标元素的 XPath", Required = true };
        var rawSetActionOpt = new Option<string>("--action") { Description = "操作：append、prepend、insertbefore、insertafter、replace、remove、setattr", Required = true };
        var rawSetXmlOpt = new Option<string?>("--xml") { Description = "XML fragment，或供 setattr 使用的 attr=value" };

        var rawSetCommand = new Command("raw-set", "修改文档 part 的原始 XML（适用于任意 OpenXML 操作的通用回退）");
        rawSetCommand.Add(rawSetFileArg);
        rawSetCommand.Add(rawSetPartArg);
        rawSetCommand.Add(rawSetXpathOpt);
        rawSetCommand.Add(rawSetActionOpt);
        rawSetCommand.Add(rawSetXmlOpt);
        rawSetCommand.Add(jsonOption);

        rawSetCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(rawSetFileArg)!;
            var partPath = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(rawSetPartArg)!)!;
            var xpath = result.GetValue(rawSetXpathOpt)!;
            var action = result.GetValue(rawSetActionOpt)!;
            var xml = result.GetValue(rawSetXmlOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "raw-set";
                req.Args["part"] = partPath;
                req.Args["xpath"] = xpath;
                req.Args["action"] = action;
                if (xml != null) req.Args["xml"] = xml;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var errorsBefore = handler.Validate().Select(e => e.Description).ToHashSet();
            handler.RawSet(partPath, xpath, action, xml);
            var warnings = ReportNewErrorsAsWarnings(handler, errorsBefore);
            var message = $"raw-set applied: {action} at {xpath}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message, warnings));
            else
            {
                Console.WriteLine(message);
                ReportNewErrors(handler, errorsBefore, warnings);
            }
            NotifyWatch(handler, file.FullName, null);
            return warnings is { Count: > 0 } ? 1 : 0;
        }, json); });

        return rawSetCommand;
    }

    private static Command BuildAddPartCommand(Option<bool> jsonOption)
    {
        var addPartFileArg = new Argument<string>("file") { Description = "文档文件路径" };
        var addPartParentArg = new Argument<string>("parent") { Description = "父 part 路径（例如文档根为 /、Excel sheet 为 /Sheet1、PPT slide 为 /slide[0]）" };
        var addPartTypeOpt = new Option<string>("--type") { Description = "要创建的 part 类型。Word：chart、header、footer；PPT/Excel：chart", Required = true };
        var addPartCommand = new Command("add-part", "创建新的文档 part，并返回供 raw-set 使用的 relationship ID");
        addPartCommand.Add(addPartFileArg);
        addPartCommand.Add(addPartParentArg);
        addPartCommand.Add(addPartTypeOpt);
        addPartCommand.Add(jsonOption);

        addPartCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(addPartFileArg)!;
            var parent = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(addPartParentArg)!)!;
            var type = result.GetValue(addPartTypeOpt)!;

            if (TryResident(file, req =>
            {
                req.Command = "add-part";
                req.Args["parent"] = parent;
                req.Args["type"] = type;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file, editable: true);
            var errorsBefore = handler.Validate().Select(e => e.Description).ToHashSet();
            var (relId, partPath) = handler.AddPart(parent, type);
            var warnings = ReportNewErrorsAsWarnings(handler, errorsBefore);
            var message = $"Created {type} part: relId={relId} path={partPath}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message, warnings));
            else
            {
                Console.WriteLine(message);
                ReportNewErrors(handler, errorsBefore, warnings);
            }
            NotifyWatch(handler, file, null);
            return 0;
        }, json); });

        return addPartCommand;
    }
}
