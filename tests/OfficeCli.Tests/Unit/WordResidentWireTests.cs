using System.IO.Pipes;
using System.Text.Json;
using OfficeCli;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class WordResidentWireTests : WordTestBase
{
    [Fact]
    public void ResidentRequestAndResponse_RoundTripAllWireFields()
    {
        var request = new ResidentRequest
        {
            Command = "set",
            Args = new() { ["path"] = "/body/p[1]", ["mode"] = "json" },
            Props = new() { ["align"] = "center" },
            Json = true
        };

        var json = JsonSerializer.Serialize(request);
        using var requestDocument = JsonDocument.Parse(json);
        var root = requestDocument.RootElement;

        Assert.Equal("set", root.GetProperty("Command").GetString());
        Assert.Equal("/body/p[1]", root.GetProperty("Args").GetProperty("path").GetString());
        Assert.Equal("center", root.GetProperty("Props").GetProperty("align").GetString());
        Assert.True(root.GetProperty("Json").GetBoolean());
        Assert.Equal("/body/p[1]", request.GetArg("path"));
        Assert.Null(request.GetArgOrNull("missing"));

        var response = new ResidentResponse
        {
            ExitCode = 1,
            Stdout = "partial output",
            Stderr = "busy"
        };
        var responseJson = JsonSerializer.Serialize(response);
        var roundTrip = JsonSerializer.Deserialize<ResidentResponse>(responseJson);

        Assert.NotNull(roundTrip);
        Assert.Equal(response.ExitCode, roundTrip!.ExitCode);
        Assert.Equal(response.Stdout, roundTrip.Stdout);
        Assert.Equal(response.Stderr, roundTrip.Stderr);
    }

    [Fact]
    public async Task TrySend_DoesNotRetryAfterConnectedReadFailure()
    {
        var path = CreateBlankDocx();
        var pipeName = ResidentServer.GetPipeName(path);
        using var server = new NamedPipeServerStream(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: 1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous);

        var acceptTask = server.WaitForConnectionAsync();
        var sendTask = Task.Run(() => ResidentClient.TrySend(
            path,
            new ResidentRequest
            {
                Command = "add",
                Args = new() { ["parent"] = "/body", ["type"] = "paragraph" },
                Props = new() { ["text"] = "must not duplicate" },
                Json = true
            },
            maxRetries: 2,
            connectTimeoutMs: 1000));

        await acceptTask.WaitAsync(TimeSpan.FromSeconds(2));
        using (var reader = new StreamReader(server, leaveOpen: true))
        {
            var requestLine = await reader.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(2));
            Assert.NotNull(requestLine);
            using var request = JsonDocument.Parse(requestLine!);
            Assert.Equal("add", request.RootElement.GetProperty("Command").GetString());
        }

        server.Dispose();
        var response = await sendTask.WaitAsync(TimeSpan.FromSeconds(2));

        Assert.Null(response);
    }

    [Fact]
    public void TrySend_WhenResidentIsMissingReturnsNullWithoutChangingTheFile()
    {
        var path = CreateBlankDocx();
        var before = File.ReadAllBytes(path);

        Assert.False(ResidentClient.TryConnect(path, out _));
        var response = ResidentClient.TrySend(
            path,
            new ResidentRequest
            {
                Command = "add",
                Args = new() { ["parent"] = "/body", ["type"] = "paragraph" },
                Props = new() { ["text"] = "must not be applied" },
                Json = true
            },
            maxRetries: 1,
            connectTimeoutMs: 50);

        Assert.Null(response);
        Assert.Equal(before, File.ReadAllBytes(path));
    }
}
