// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildWatchCommand(Option<bool> jsonOption)
    {
        var watchFileArg = new Argument<FileInfo>("file") { Description = "Office 文档路径（.pptx、.xlsx、.docx）" };
        var watchPortOpt = new Option<int>("--port") { Description = "预览 server 的 HTTP port" };
        watchPortOpt.DefaultValueFactory = _ => 26315;

        var watchCommand = new Command("watch", "启动实时预览 server；officecli 修改文档时自动刷新（不检测外部编辑）。子命令 mark/unmark/marks/goto 操作正在运行的预览。");
        watchCommand.Add(watchFileArg);
        watchCommand.Add(watchPortOpt);

        // Subcommands — operate against the running watch process via named-pipe IPC.
        // These were previously top-level (`mark`, `unmark`, `get-marks`, `goto`);
        // grouped under `watch` to reflect that they only function while a watch
        // session is alive. The top-level forms remain registered as hidden BC
        // aliases (see CommandBuilder.cs).
        watchCommand.Add(BuildMarkCommand(jsonOption, "mark"));
        watchCommand.Add(BuildUnmarkMarkCommand(jsonOption, "unmark"));
        watchCommand.Add(BuildGetMarksCommand(jsonOption, "marks"));
        watchCommand.Add(BuildGotoCommand(jsonOption, "goto"));

        watchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(watchFileArg)!;
            var port = result.GetValue(watchPortOpt);

            // Render initial HTML: ask the resident process if one is running,
            // otherwise open the file directly as a fallback.
            string? initialHtml = null;
            if (file.Exists)
            {
                // Try resident first — avoids file lock conflict.
                // Json=true makes resident return raw HTML via Console.Write;
                // the resident then wraps it in a JSON envelope { "success":true, "message":"<html>..." }.
                var resp = ResidentClient.TrySend(file.FullName,
                    new ResidentRequest { Command = "view", Args = new() { ["mode"] = "html" }, Json = true },
                    connectTimeoutMs: 2000);
                if (resp is { ExitCode: 0 } && !string.IsNullOrEmpty(resp.Stdout))
                {
                    try
                    {
                        using var doc = System.Text.Json.JsonDocument.Parse(resp.Stdout);
                        if (doc.RootElement.TryGetProperty("message", out var msg))
                            initialHtml = msg.GetString();
                    }
                    catch { /* parse failed — fall through to direct open */ }
                }
                else
                {
                    // No resident — open directly
                    try
                    {
                        using var handler = DocumentHandlerFactory.Open(file.FullName, editable: false);
                        if (handler is OfficeCli.Handlers.PowerPointHandler)
                            initialHtml = RenderViaRegistry(handler, "pptx", new OfficeCli.Core.Rendering.RenderOptions());
                        else if (handler is OfficeCli.Handlers.ExcelHandler)
                            initialHtml = RenderViaRegistry(handler, "xlsx", new OfficeCli.Core.Rendering.RenderOptions());
                        else if (handler is OfficeCli.Handlers.WordHandler)
                            initialHtml = RenderViaRegistry(handler, "docx", new OfficeCli.Core.Rendering.RenderOptions());
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Warning: initial render failed — preview will show 'Waiting for first update' until the next document change.");
                        Console.Error.WriteLine($"  {ex.GetType().Name}: {ex.Message}");
                        if (Environment.GetEnvironmentVariable("OFFICECLI_DEBUG") == "1" && ex.StackTrace != null)
                            Console.Error.WriteLine(ex.StackTrace);
                    }
                }
            }

            using var cts = new CancellationTokenSource();

            using var watch = new WatchServer(file.FullName, port, initialHtml: initialHtml);
            // Signal handling (SIGTERM / SIGINT / SIGHUP / SIGQUIT) is
            // now registered inside WatchServer.RunAsync via
            // PosixSignalRegistration, which runs BEFORE the .NET runtime
            // begins its shutdown sequence (on a healthy ThreadPool).
            // That path runs StopAsync to completion — including
            // TcpListener.Stop() (the only reliable way to unstick
            // AcceptTcpClientAsync on macOS) and the CoreFxPipe_ socket
            // cleanup (BUG-BT-003) — before calling Environment.Exit.
            //
            // The older Console.CancelKeyPress + ProcessExit combo was
            // unreliable: SIGINT would cancel _cts but the TCP accept
            // loop did not honour cancellation on macOS, hanging the
            // process for 15+ seconds; ProcessExit ran during runtime
            // teardown when ThreadPool was already unwinding, so the
            // socket cleanup silently skipped.
            watch.RunAsync(cts.Token).GetAwaiter().GetResult();
            return 0;
        }));

        return watchCommand;
    }

    private static Command BuildUnwatchCommand()
    {
        var unwatchFileArg = new Argument<FileInfo>("file") { Description = "Office 文档路径（.pptx、.xlsx、.docx）" };
        var unwatchCommand = new Command("unwatch", "停止该文档的 watch 预览 server");
        unwatchCommand.Add(unwatchFileArg);

        unwatchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(unwatchFileArg)!;
            if (WatchNotifier.SendClose(file.FullName))
                Console.WriteLine($"Watch stopped for {file.Name}");
            else
                Console.Error.WriteLine($"No watch running for {file.Name}");
            return 0;
        }));

        return unwatchCommand;
    }
}
