using System.Net;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli.Tests.Integration;

[Collection("Word CLI output")]
[Trait("Speed", "Integration")]
public sealed class WordWatchServerIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public async Task WatchServer_ServesHtmlSseAndSelection_AndCleansMarkerOnStop()
    {
        var path = CreateBlankDocx();
        var port = GetFreePort();
        using var server = new WatchServer(
            path,
            port,
            idleTimeout: TimeSpan.FromSeconds(30),
            initialHtml: "<html><body>watch fixture</body></html>");
        var serverTask = server.RunAsync();

        try
        {
            WaitUntil(() => WatchServer.IsWatching(path), "watch marker was not created");

            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };
            var page = await client.GetAsync($"http://127.0.0.1:{port}/");
            Assert.Equal(HttpStatusCode.OK, page.StatusCode);
            Assert.Equal("text/html", page.Content.Headers.ContentType?.MediaType);
            Assert.Contains("watch fixture", await page.Content.ReadAsStringAsync());

            using var sseResponse = await client.GetAsync(
                $"http://127.0.0.1:{port}/events",
                HttpCompletionOption.ResponseHeadersRead);
            Assert.Equal(HttpStatusCode.OK, sseResponse.StatusCode);
            Assert.Equal("text/event-stream", sseResponse.Content.Headers.ContentType?.MediaType);
            using var sseStream = await sseResponse.Content.ReadAsStreamAsync();
            using (var readTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(2)))
            {
                var buffer = new byte[2048];
                var read = await sseStream.ReadAsync(buffer, readTimeout.Token);
                Assert.Contains("selection-update", Encoding.UTF8.GetString(buffer, 0, read));
            }

            using var selectionContent = new StringContent(
                "{\"paths\":[\"/body/p[1]\"]}", Encoding.UTF8, "application/json");
            var selection = await client.PostAsync(
                $"http://127.0.0.1:{port}/api/selection", selectionContent);
            Assert.Equal(HttpStatusCode.NoContent, selection.StatusCode);
            Assert.Equal(new[] { "/body/p[1]" }, WatchNotifier.QuerySelection(path));

            var methodError = await client.GetAsync($"http://127.0.0.1:{port}/api/selection");
            Assert.Equal(HttpStatusCode.MethodNotAllowed, methodError.StatusCode);
            var unknownApi = await client.GetAsync($"http://127.0.0.1:{port}/api/unknown");
            Assert.Equal(HttpStatusCode.NotFound, unknownApi.StatusCode);

            using var hostile = new HttpRequestMessage(
                HttpMethod.Get, $"http://127.0.0.1:{port}/");
            hostile.Headers.Host = "evil.example";
            var forbidden = await client.SendAsync(hostile);
            Assert.Equal(HttpStatusCode.Forbidden, forbidden.StatusCode);
        }
        finally
        {
            await server.StopAsync();
            await serverTask.WaitAsync(TimeSpan.FromSeconds(5));
        }

        Assert.False(WatchServer.IsWatching(path));
        Assert.Null(WatchServer.GetExistingWatchPort(path));
    }

    private static int GetFreePort()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
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
