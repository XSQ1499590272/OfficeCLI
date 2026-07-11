using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordResidentConcurrencyTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public async Task ConcurrentAdds_AreSerializedAndAppliedExactlyOnce()
    {
        var path = CreateBlankDocx();
        using var server = new ResidentServer(path);
        var serverTask = Task.Run(() => server.RunAsync());

        try
        {
            WaitUntil(() => ResidentClient.TryConnect(path, out _), "resident did not start");

            var responses = await Task.WhenAll(
                Enumerable.Range(0, 8).Select(index => Task.Run(() =>
                    ResidentClient.TrySend(path, new ResidentRequest
                    {
                        Command = "add",
                        Args = new() { ["parent"] = "/body", ["type"] = "paragraph" },
                        Props = new() { ["text"] = $"concurrent-{index}" },
                        Json = true
                    }, maxRetries: 4, connectTimeoutMs: 1000))));

            Assert.All(responses, response =>
            {
                Assert.NotNull(response);
                Assert.Equal(0, response!.ExitCode);
            });

            var query = ResidentClient.TrySend(path, new ResidentRequest
            {
                Command = "query",
                Args = new() { ["selector"] = "paragraph" },
                Json = true
            });

            Assert.NotNull(query);
            Assert.Equal(0, query!.ExitCode);
            foreach (var index in Enumerable.Range(0, 8))
                Assert.True(CountTopLevelResultTexts(query.Stdout, $"concurrent-{index}") == 1, query.Stdout);

            Assert.True(ResidentClient.SendSave(path));
            using (var reopened = new WordHandler(path, editable: false))
            {
                var texts = reopened.Query("paragraph").Select(node => node.Text).ToList();
                Assert.Equal(8, texts.Count);
                Assert.Equal(Enumerable.Range(0, 8).Select(index => $"concurrent-{index}").OrderBy(text => text), texts.OrderBy(text => text));
            }

            Assert.True(ResidentClient.SendClose(path));
            WaitUntil(() => !ResidentClient.TryConnect(path, out _), "resident did not close");
            await serverTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
        finally
        {
            try { ResidentClient.SendClose(path); } catch { }
        }
    }

    private static int CountTopLevelResultTexts(string json, string expectedText)
    {
        using var document = System.Text.Json.JsonDocument.Parse(json);
        var root = document.RootElement;
        if (root.TryGetProperty("data", out var data)) root = data;
        var results = root.GetProperty("results");
        return results.EnumerateArray().Count(item =>
            item.TryGetProperty("text", out var text)
            && string.Equals(text.GetString(), expectedText, StringComparison.Ordinal));
    }

    private static void WaitUntil(Func<bool> condition, string message)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(5);
        while (DateTime.UtcNow < deadline)
        {
            if (condition()) return;
            Thread.Sleep(50);
        }

        Assert.Fail(message);
    }
}
