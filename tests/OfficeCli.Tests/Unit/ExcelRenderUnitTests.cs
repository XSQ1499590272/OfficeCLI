// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Unit;

[CollectionDefinition(Name, DisableParallelization = true)]
[Trait("Speed", "Unit")]
public sealed class ExcelRenderBackendCollectionDefinition
{
    public const string Name = "Excel render backend";
}

[Collection(ExcelRenderBackendCollectionDefinition.Name)]
public sealed class ExcelRenderUnitTests : ExcelTestBase
{
    private static readonly object PathMutationLock = new();

    [Fact]
    public void HtmlScreenshot_AutoGridColumns_TracksCountAndAspectRatio()
    {
        HtmlScreenshot.AutoGridColumns(1, 160, 90).Should().Be(1);
        HtmlScreenshot.AutoGridColumns(4, 200, 100).Should().Be(1);
        HtmlScreenshot.AutoGridColumns(4, 100, 100).Should().Be(2);
        HtmlScreenshot.AutoGridColumns(9, 100, 200).Should().Be(4);
    }

    [Fact]
    public void HtmlScreenshot_Capture_UsesSuccessfulPlaywrightBackend()
    {
        var toolDir = TrackTempDirectory();
        var htmlPath = TrackTempFile(Path.Combine(toolDir, "capture.html"));
        var outPath = TrackTempFile(Path.Combine(toolDir, "capture.png"));
        File.WriteAllText(htmlPath, "<html><body>capture</body></html>");
        WritePlaywrightSuccessExecutable(toolDir);

        using var _ = PrependPath(toolDir);

        var result = HtmlScreenshot.Capture(htmlPath, outPath, width: 400, height: 300);
        result.Ok.Should().BeTrue();
        result.Backend.Should().Be("playwright");
        result.Error.Should().BeNull();
        File.ReadAllText(outPath).Should().Be("png");
    }

    [Fact]
    public void HtmlScreenshot_Capture_FallsBackAfterBackendFailure()
    {
        var toolDir = TrackTempDirectory();
        var htmlPath = TrackTempFile(Path.Combine(toolDir, "capture.html"));
        var outPath = TrackTempFile(Path.Combine(toolDir, "capture.png"));
        File.WriteAllText(htmlPath, "<html><body>capture</body></html>");
        WritePlaywrightFailureExecutable(toolDir);
        WriteChromeSuccessExecutable(toolDir);

        using var _ = PrependPath(toolDir);

        var result = HtmlScreenshot.Capture(htmlPath, outPath, width: 400, height: 300);
        result.Ok.Should().BeTrue();
        result.Backend.Should().Be("chrome");
        result.Error.Should().BeNull();
        File.ReadAllText(outPath).Should().Be("png");
    }

    [Fact]
    public void HtmlScreenshot_Capture_ReturnsNoBackendFallbackWhenBackendProducesEmptyOutput()
    {
        var toolDir = TrackTempDirectory();
        var htmlPath = TrackTempFile(Path.Combine(toolDir, "capture.html"));
        var outPath = TrackTempFile(Path.Combine(toolDir, "capture.png"));
        File.WriteAllText(htmlPath, "<html><body>capture</body></html>");
        WritePlaywrightEmptyExecutable(toolDir);
        WriteChromeEmptyExecutable(toolDir);
        WriteFirefoxEmptyExecutable(toolDir);

        using var _ = PrependPath(toolDir);

        var result = HtmlScreenshot.Capture(htmlPath, outPath, width: 400, height: 300);
        result.Ok.Should().BeFalse();
        result.Backend.Should().BeEmpty();
        result.Error.Should().Be("no headless backend available");
        new FileInfo(outPath).Length.Should().Be(0);
    }

    [Fact]
    public void ViewAsHtml_AppliesNormalizedCellColorsAndNumberFormatSectionColors()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "Alert"), ("font.color", "ff0000")));
            handler.Add("/Sheet1/B1", "cell", null, Props(("value", "-5"), ("numberformat", "0;[Red](0)")));
            handler.Add("/Sheet1/C1", "cell", null, Props(("value", "BlueText"), ("numberformat", "0;0;0;[Blue]@")));
            handler.Add("/Sheet1/D1", "cell", null, Props(("value", "60"), ("numberformat", "[Green][>50]0;[Red]0")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var html = readOnly.ViewAsHtml();

        html.Should().Contain("#FF0000");
        html.Should().Contain("Alert");
        html.Should().Contain("(5)");
        html.Should().Contain("#0000FF");
        html.Should().Contain("BlueText");
        html.Should().Contain("#00FF00");
        html.Should().Contain(">60<");
    }

    private static Dictionary<string, string> Props(params (string Key, string Value)[] items)
        => items.ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase);

    private static IDisposable PrependPath(string toolDir)
    {
        System.Threading.Monitor.Enter(PathMutationLock);
        var originalPath = Environment.GetEnvironmentVariable("PATH");
        Environment.SetEnvironmentVariable("PATH", $"{toolDir}{Path.PathSeparator}{originalPath}");
        return new RestorePathDisposable(originalPath);
    }

    private static void WritePlaywrightSuccessExecutable(string toolDir)
    {
        if (OperatingSystem.IsWindows())
        {
            WriteWindowsScript(toolDir, "playwright", """
                @echo off
                set out=%5
                >"%out%" echo|set /p =png
                """);
            return;
        }

        WriteUnixScript(toolDir, "playwright", """
            out=""
            for last do
              out="$last"
            done
            printf 'png' > "$out"
            """);
    }

    private static void WritePlaywrightFailureExecutable(string toolDir)
    {
        if (OperatingSystem.IsWindows())
        {
            WriteWindowsScript(toolDir, "playwright", """
                @echo off
                >&2 echo playwright boom
                exit /b 7
                """);
            return;
        }

        WriteUnixScript(toolDir, "playwright", """
            echo playwright boom >&2
            exit 7
            """);
    }

    private static void WritePlaywrightEmptyExecutable(string toolDir)
    {
        if (OperatingSystem.IsWindows())
        {
            WriteWindowsScript(toolDir, "playwright", """
                @echo off
                set out=%5
                type nul > "%out%"
                """);
            return;
        }

        WriteUnixScript(toolDir, "playwright", """
            out=""
            for last do
              out="$last"
            done
            : > "$out"
            """);
    }

    private static void WriteChromeSuccessExecutable(string toolDir)
    {
        if (OperatingSystem.IsWindows())
        {
            WriteWindowsScript(toolDir, "chrome", """
                @echo off
                set out=
                for %%a in (%*) do (
                  set "arg=%%~a"
                  call set "probe=%%arg:~0,13%%"
                  if /I "%%probe%%"=="--screenshot=" call set "out=%%arg:~13%%"
                )
                >"%out%" echo|set /p =png
                """);
            return;
        }

        WriteUnixScript(toolDir, "chrome", """
            out=""
            for arg in "$@"; do
              case "$arg" in
                --screenshot=*) out="${arg#--screenshot=}" ;;
              esac
            done
            printf 'png' > "$out"
            """);
    }

    private static void WriteChromeEmptyExecutable(string toolDir)
    {
        if (OperatingSystem.IsWindows())
        {
            WriteWindowsScript(toolDir, "chrome", """
                @echo off
                set out=
                for %%a in (%*) do (
                  set "arg=%%~a"
                  call set "probe=%%arg:~0,13%%"
                  if /I "%%probe%%"=="--screenshot=" call set "out=%%arg:~13%%"
                )
                type nul > "%out%"
                """);
            return;
        }

        WriteUnixScript(toolDir, "chrome", """
            out=""
            for arg in "$@"; do
              case "$arg" in
                --screenshot=*) out="${arg#--screenshot=}" ;;
              esac
            done
            : > "$out"
            """);
    }

    private static void WriteFirefoxEmptyExecutable(string toolDir)
    {
        if (OperatingSystem.IsWindows())
        {
            WriteWindowsScript(toolDir, "firefox", """
                @echo off
                set out=
                for %%a in (%*) do (
                  set "arg=%%~a"
                  call set "probe=%%arg:~0,13%%"
                  if /I "%%probe%%"=="--screenshot=" call set "out=%%arg:~13%%"
                )
                type nul > "%out%"
                """);
            return;
        }

        WriteUnixScript(toolDir, "firefox", """
            out=""
            for arg in "$@"; do
              case "$arg" in
                --screenshot=*) out="${arg#--screenshot=}" ;;
              esac
            done
            : > "$out"
            """);
    }

    private static void WriteWindowsScript(string toolDir, string baseName, string content)
    {
        var path = Path.Combine(toolDir, $"{baseName}.cmd");
        File.WriteAllText(path, content.Replace("\n", "\r\n"));
    }

    private static void WriteUnixScript(string toolDir, string baseName, string content)
    {
        var path = Path.Combine(toolDir, baseName);
        File.WriteAllText(path, "#!/bin/sh\n" + content);
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
    }

    private sealed class RestorePathDisposable(string? originalPath) : IDisposable
    {
        public void Dispose()
        {
            Environment.SetEnvironmentVariable("PATH", originalPath);
            System.Threading.Monitor.Exit(PathMutationLock);
        }
    }
}
