using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using OfficeCli;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordResidentFailureIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    private static readonly TimeSpan AsyncTimeout = TimeSpan.FromSeconds(10);

    [Fact]
    public async Task BusyMainPipe_RetryIsLimitedToConnectAndDoesNotChangeDocument()
    {
        var path = CreateBlankDocx();
        var pipeName = ResidentServer.GetPipeName(path);
        using var server = new NamedPipeServerStream(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: 1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous);
        using var occupiedClient = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);

        var acceptTask = server.WaitForConnectionAsync();
        occupiedClient.Connect(1000);
        await acceptTask.WaitAsync(AsyncTimeout);

        var before = File.ReadAllBytes(path);
        var sendTask = Task.Run(() => ResidentClient.TrySend(
            path,
            new ResidentRequest
            {
                Command = "add",
                Args = new() { ["parent"] = "/body", ["type"] = "paragraph" },
                Props = new() { ["text"] = "must not be duplicated" },
                Json = true
            },
            maxRetries: 2,
            connectTimeoutMs: 50));

        // macOS can keep a second named-pipe connect pending when the only
        // server instance is occupied. Release the fixture if the platform
        // does not honor the short connect timeout, then observe the same
        // null/no-mutation contract without allowing the test to hang.
        if (await Task.WhenAny(sendTask, Task.Delay(1000)) != sendTask)
        {
            server.Dispose();
            occupiedClient.Dispose();
        }

        Assert.Null(await sendTask.WaitAsync(AsyncTimeout));
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Fact]
    public async Task PingForDifferentFile_IsRejectedAsStaleResident()
    {
        var path = CreateBlankDocx();
        var otherPath = CreateBlankDocx();
        var pipeName = ResidentServer.GetPipeName(path) + "-ping";
        using var server = new NamedPipeServerStream(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: 1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous);

        var acceptTask = server.WaitForConnectionAsync();
        var pingTask = Task.Run(() => ResidentClient.TryConnect(path, out _));
        await acceptTask.WaitAsync(AsyncTimeout);

        using (var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true))
        using (var writer = new StreamWriter(server, new UTF8Encoding(false), leaveOpen: true)
        { AutoFlush = true })
        {
            Assert.NotNull(await reader.ReadLineAsync().WaitAsync(AsyncTimeout));
            var response = new ResidentResponse { ExitCode = 0, Stdout = Path.GetFullPath(otherPath) };
            await writer.WriteLineAsync(JsonSerializer.Serialize(response));
        }

        Assert.False(await pingTask.WaitAsync(AsyncTimeout));
    }

    [Fact]
    public async Task ConnectedHalfResponse_IsNotRetriedAfterCommandMayHaveApplied()
    {
        var path = CreateBlankDocx();
        var pipeName = ResidentServer.GetPipeName(path);
        using var server = new NamedPipeServerStream(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: 1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous);

        var requestCount = 0;
        var acceptTask = server.WaitForConnectionAsync();
        var sendTask = Task.Run(() => ResidentClient.TrySend(
            path,
            new ResidentRequest
            {
                Command = "add",
                Args = new() { ["parent"] = "/body", ["type"] = "paragraph" },
                Props = new() { ["text"] = "at-most-once" },
                Json = true
            },
            maxRetries: 3,
            connectTimeoutMs: 1000));

        await acceptTask.WaitAsync(AsyncTimeout);
        using (var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true))
        {
            Assert.NotNull(await reader.ReadLineAsync().WaitAsync(AsyncTimeout));
            requestCount++;
        }

        var partialResponse = Encoding.UTF8.GetBytes("{\"ExitCode\":0");
        await server.WriteAsync(partialResponse);
        await server.FlushAsync();
        server.Dispose();

        Assert.Null(await sendTask.WaitAsync(AsyncTimeout));
        Assert.Equal(1, requestCount);
    }

    [Fact]
    public void ResidentProcessExit_LeavesNoReachablePipeAndForceCreateRecovers()
    {
        var path = CreateBlankDocx();
        var appHost = FindOfficeCliAppHost();
        var startInfo = new System.Diagnostics.ProcessStartInfo(appHost)
        {
            WorkingDirectory = Path.GetDirectoryName(appHost)!,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "off";

        using var resident = System.Diagnostics.Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start resident");
        _ = resident.StandardOutput.ReadToEndAsync();
        _ = resident.StandardError.ReadToEndAsync();

        try
        {
            WaitUntil(() => ResidentClient.TryConnect(path, out _), "resident did not start");
            resident.Kill(entireProcessTree: true);
            Assert.True(resident.WaitForExit(5_000), "resident did not exit after kill");
            WaitUntil(() => !ResidentClient.TryConnect(path, out _), "resident pipe remained reachable");

            var recovered = RunOfficeCli("create", path, "--force", "--json");
            Assert.Equal(0, recovered.ExitCode);
            using var json = JsonDocument.Parse(recovered.Stdout);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean(), recovered.Stdout);
        }
        finally
        {
            if (!resident.HasExited)
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
                try { resident.WaitForExit(5_000); } catch { }
            }
        }
    }

    private static string FindOfficeCliAppHost()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "officecli.slnx")))
            directory = directory.Parent;

        var root = directory?.FullName
            ?? throw new InvalidOperationException("Could not locate officecli.slnx from test output directory.");
        var executableName = OperatingSystem.IsWindows() ? "officecli.exe" : "officecli";
        return Directory.GetFiles(
                Path.Combine(root, "src", "officecli", "bin"),
                executableName,
                SearchOption.AllDirectories)
            .Where(file => new FileInfo(file).Length > 0)
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .FirstOrDefault()
            ?? throw new InvalidOperationException("Could not locate built officecli apphost.");
    }

    private static ProcessResult RunOfficeCli(params string[] args)
    {
        var appHost = FindOfficeCliAppHost();
        var startInfo = new System.Diagnostics.ProcessStartInfo(appHost)
        {
            WorkingDirectory = Path.GetDirectoryName(appHost)!,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        foreach (var arg in args)
            startInfo.ArgumentList.Add(arg);
        startInfo.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";

        using var process = System.Diagnostics.Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start officecli");
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        Assert.True(process.WaitForExit(10_000), "officecli did not exit");
        return new ProcessResult(process.ExitCode, stdout, stderr);
    }

    private static void WaitUntil(Func<bool> condition, string message)
    {
        var deadline = DateTime.UtcNow.Add(AsyncTimeout);
        while (DateTime.UtcNow < deadline)
        {
            if (condition()) return;
            Thread.Sleep(50);
        }
        throw new TimeoutException(message);
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
