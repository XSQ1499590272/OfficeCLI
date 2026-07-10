using System.Text.Json;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public class WordResidentLifecycleTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public async Task ResidentSaveAndCloseControlDiskVisibility()
    {
        var previousFlush = Environment.GetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH");
        Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH", "off");
        var path = CreateBlankDocx();

        using var server = new ResidentServer(path);
        var serverTask = Task.Run(() => server.RunAsync());
        try
        {
            WaitUntil(() => ResidentClient.TryConnect(path, out _), "resident did not start");

            var add = ResidentClient.TrySend(path, new ResidentRequest
            {
                Command = "add",
                Args = { ["parent"] = "/body", ["type"] = "paragraph" },
                Props = new() { ["text"] = "resident pending" },
                Json = true
            });
            Assert.NotNull(add);
            Assert.Equal(0, add.ExitCode);

            var query = ResidentClient.TrySend(path, new ResidentRequest
            {
                Command = "query",
                Args = { ["selector"] = "paragraph", ["find"] = "resident pending" },
                Json = true
            });
            Assert.NotNull(query);
            Assert.Equal(0, query.ExitCode);
            Assert.Contains("resident pending", query.Stdout);

            using (var beforeSave = new WordHandler(path, editable: false))
            {
                Assert.DoesNotContain("resident pending", beforeSave.Query("paragraph").Select(p => p.Text));
            }

            Assert.True(ResidentClient.SendSave(path));
            using (var afterSave = new WordHandler(path, editable: false))
            {
                Assert.Contains("resident pending", afterSave.Query("paragraph").Select(p => p.Text));
            }

            var set = ResidentClient.TrySend(path, new ResidentRequest
            {
                Command = "set",
                Args = { ["path"] = "/body/p[1]" },
                Props = new() { ["align"] = "right" },
                Json = true
            });
            Assert.NotNull(set);
            Assert.Equal(0, set.ExitCode);

            Assert.True(ResidentClient.SendClose(path));
            WaitUntil(() => !ResidentClient.TryConnect(path, out _), "resident did not close");
            await serverTask.WaitAsync(TimeSpan.FromSeconds(5));

            using var afterClose = new WordHandler(path, editable: false);
            Assert.Equal("right", afterClose.Get("/body/p[1]").Format["align"]);
        }
        finally
        {
            try { ResidentClient.SendClose(path); } catch { }
            Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH", previousFlush);
        }
    }

    [Fact]
    public async Task ResidentBatchSkipsOpenCloseAndDefersDiskFlushUntilSave()
    {
        var previousFlush = Environment.GetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH");
        Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH", "off");
        var path = CreateBlankDocx();

        using var server = new ResidentServer(path);
        var serverTask = Task.Run(() => server.RunAsync());
        try
        {
            WaitUntil(() => ResidentClient.TryConnect(path, out _), "resident did not start");

            var batch = ResidentClient.TrySend(path, new ResidentRequest
            {
                Command = "batch",
                Args =
                {
                    ["force"] = "true",
                    ["batchJson"] = """
                        [
                          {"command":"open"},
                          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"resident batch"}},
                          {"command":"close"},
                          {"command":"query","selector":"paragraph","text":"resident batch"}
                        ]
                        """
                },
                Json = true
            });

            Assert.NotNull(batch);
            Assert.Equal(0, batch.ExitCode);
            Assert.Contains("Skipped 'open' (resident mode)", batch.Stdout);
            Assert.Contains("Added paragraph", batch.Stdout);
            Assert.Contains("Skipped 'close' (resident mode)", batch.Stdout);
            Assert.Contains("resident batch", batch.Stdout);
            Assert.True(ResidentClient.TryConnect(path, out _));

            using (var beforeSave = new WordHandler(path, editable: false))
            {
                Assert.DoesNotContain("resident batch", beforeSave.Query("paragraph").Select(p => p.Text));
            }

            Assert.True(ResidentClient.SendSave(path));
            using (var afterSave = new WordHandler(path, editable: false))
            {
                Assert.Contains("resident batch", afterSave.Query("paragraph").Select(p => p.Text));
            }

            Assert.True(ResidentClient.SendClose(path));
            WaitUntil(() => !ResidentClient.TryConnect(path, out _), "resident did not close");
            await serverTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
        finally
        {
            try { ResidentClient.SendClose(path); } catch { }
            Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH", previousFlush);
        }
    }

    [Fact]
    public async Task ResidentBatchJsonEnvelopeMatchesNonResidentResultsAndSummary()
    {
        const string itemsJson = """
        [
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"resident parity"}},
          {"command":"add","parent":"/body/p[99]","type":"paragraph","props":{"text":"should fail"}},
          {"command":"query","selector":"paragraph"}
        ]
        """;

        var nonResidentPath = CreateBlankDocx();
        using (var handler = new WordHandler(nonResidentPath, editable: true))
        {
            var nonResident = BatchExecutor.ExecuteBatch(handler, itemsJson, json: true);
            AssertBatchEnvelope(nonResident, total: 3, executed: 3, succeeded: 2, failed: 1, skipped: 0);
        }

        var previousFlush = Environment.GetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH");
        Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH", "off");
        var residentPath = CreateBlankDocx();

        using var server = new ResidentServer(residentPath);
        var serverTask = Task.Run(() => server.RunAsync());
        try
        {
            WaitUntil(() => ResidentClient.TryConnect(residentPath, out _), "resident did not start");

            var response = ResidentClient.TrySend(residentPath, new ResidentRequest
            {
                Command = "batch",
                Args =
                {
                    ["batchJson"] = itemsJson,
                    ["force"] = "true",
                    ["stopOnError"] = "false"
                },
                Json = true
            });

            Assert.NotNull(response);
            Assert.Equal(1, response.ExitCode);
            AssertBatchEnvelope(response.Stdout, total: 3, executed: 3, succeeded: 2, failed: 1, skipped: 0);

            Assert.True(ResidentClient.SendClose(residentPath));
            WaitUntil(() => !ResidentClient.TryConnect(residentPath, out _), "resident did not close");
            await serverTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
        finally
        {
            try { ResidentClient.SendClose(residentPath); } catch { }
            Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH", previousFlush);
        }
    }

    [Fact]
    public async Task ResidentBatchJsonEnvelope_StopOnErrorMatchesNonResidentResults()
    {
        const string itemsJson = """
        [
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"before stop"}},
          {"command":"add","parent":"/body/p[99]","type":"paragraph","props":{"text":"stop here"}},
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"after stop"}}
        ]
        """;

        var nonResidentPath = CreateBlankDocx();
        using (var handler = new WordHandler(nonResidentPath, editable: true))
        {
            var nonResident = BatchExecutor.ExecuteBatch(handler, itemsJson, json: true, stopOnError: true);
            AssertBatchEnvelope(nonResident, total: 3, executed: 2, succeeded: 1, failed: 1, skipped: 1);
        }

        var previousFlush = Environment.GetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH");
        Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH", "off");
        var residentPath = CreateBlankDocx();

        using var server = new ResidentServer(residentPath);
        var serverTask = Task.Run(() => server.RunAsync());
        try
        {
            WaitUntil(() => ResidentClient.TryConnect(residentPath, out _), "resident did not start");

            var response = ResidentClient.TrySend(residentPath, new ResidentRequest
            {
                Command = "batch",
                Args =
                {
                    ["batchJson"] = itemsJson,
                    ["force"] = "false",
                    ["stopOnError"] = "true"
                },
                Json = true
            });

            Assert.NotNull(response);
            Assert.Equal(1, response.ExitCode);
            AssertBatchEnvelope(response.Stdout, total: 3, executed: 2, succeeded: 1, failed: 1, skipped: 1);

            Assert.True(ResidentClient.SendClose(residentPath));
            WaitUntil(() => !ResidentClient.TryConnect(residentPath, out _), "resident did not close");
            await serverTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
        finally
        {
            try { ResidentClient.SendClose(residentPath); } catch { }
            Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_FLUSH", previousFlush);
        }
    }

    private static void AssertBatchEnvelope(
        string json, int total, int executed, int succeeded, int failed, int skipped)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean(), json);

        var data = root.GetProperty("data");
        var results = data.GetProperty("results");
        Assert.Equal(executed, results.GetArrayLength());
        Assert.False(results[1].GetProperty("success").GetBoolean());
        Assert.True(results[1].TryGetProperty("error", out _));
        Assert.Equal("add", results[1].GetProperty("item").GetProperty("command").GetString());

        var summary = data.GetProperty("summary");
        Assert.Equal(total, summary.GetProperty("total").GetInt32());
        Assert.Equal(executed, summary.GetProperty("executed").GetInt32());
        Assert.Equal(succeeded, summary.GetProperty("succeeded").GetInt32());
        Assert.Equal(failed, summary.GetProperty("failed").GetInt32());
        Assert.Equal(skipped, summary.GetProperty("skipped").GetInt32());
    }

    private static void WaitUntil(Func<bool> condition, string message)
    {
        var deadline = DateTime.UtcNow.AddSeconds(5);
        while (DateTime.UtcNow < deadline)
        {
            if (condition()) return;
            Thread.Sleep(50);
        }
        throw new TimeoutException(message);
    }
}
