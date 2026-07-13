// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    // ==================== save command ====================
    //
    // Flush the resident's in-memory document to disk WITHOUT ending the
    // session. Only meaningful inside a running resident — agents that build
    // a workbook incrementally and need mid-build snapshots (e.g. for a
    // third-party Excel viewer that ingests the .xlsx package directly) use
    // this to recover the parse-amortization benefit of resident mode while
    // still publishing snapshots on demand.
    //
    // Non-resident mode is rejected on purpose: each non-resident command
    // already opens-mutates-closes (close = save), so there's no pending
    // in-memory state to flush. A no-op success would invite confused
    // "save just in case" code; an error tells the user they're in the
    // wrong mode.
    private static Command BuildSaveCommand(Option<bool> jsonOption)
    {
        var saveFileArg = new Argument<FileInfo>("file") { Description = "Office 文档路径" };
        var saveCommand = new Command("save", "将内存修改写入磁盘，同时保持 resident 运行。非 officecli 程序读取文件前请执行；officecli 自身读取始终可见修改，直接磁盘读取在 flush 前只会看到修改前文件。运行中的 resident 空闲后也会自动写入（自适应 2–10 秒；参见 OFFICECLI_RESIDENT_FLUSH: each|auto|<seconds>|off）。没有活动 resident 时为 no-op。");
        saveCommand.Add(saveFileArg);
        saveCommand.Add(jsonOption);

        saveCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(saveFileArg)!;
            var filePath = file.FullName;

            // Probe without auto-starting. `save` is meaningless without a
            // pre-existing in-memory session, so we deliberately skip the
            // TryResident auto-start path that other verbs use.
            if (!ResidentClient.TryConnect(filePath, out _))
            {
                // No resident session to flush. In the non-resident model the
                // document on disk is already current (each mutation eager-saved),
                // so save is a no-op SUCCESS rather than an error — keeping
                // "edit, then save/close" a safe habit regardless of backend.
                var msg = $"{file.Name} is already saved to disk.";
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else
                    Console.WriteLine(msg);
                return 0;
            }

            var request = new ResidentRequest { Command = "save", Json = json };
            var response = ResidentClient.TrySend(
                filePath, request,
                maxRetries: ResidentBusyMaxRetries,
                connectTimeoutMs: ResidentBusyConnectTimeoutMs);
            if (response == null)
            {
                var msg = $"Resident for {file.Name} is running but the save command could not be delivered (main pipe busy or unresponsive).";
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg));
                else
                    Console.Error.WriteLine($"Error: {msg}");
                return 3;
            }

            if (!string.IsNullOrEmpty(response.Stdout))
                Console.WriteLine(response.Stdout);
            if (!string.IsNullOrEmpty(response.Stderr))
                Console.Error.WriteLine(response.Stderr);
            return response.ExitCode;
        }, json); });

        return saveCommand;
    }
}
