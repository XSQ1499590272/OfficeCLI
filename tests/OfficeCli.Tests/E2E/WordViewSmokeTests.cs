using System.Diagnostics;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public class WordViewSmokeTests : OfficeCli.Tests.Unit.WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Theory]
    [InlineData("text")]
    [InlineData("annotated")]
    [InlineData("outline")]
    [InlineData("stats")]
    [InlineData("issues")]
    public void ViewJsonModes_ReturnParseableEnvelopesWithStableKeys(string mode)
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "CLI view paragraph" });
        }

        var result = RunOfficeCli("view", path, mode, "--json");

        Assert.True(result.ExitCode == 0, $"exit={result.ExitCode}\nstdout:\n{result.Stdout}\nstderr:\n{result.Stderr}");
        using var doc = JsonDocument.Parse(result.Stdout);
        var root = doc.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), result.Stdout);
        var data = root.GetProperty("data");

        switch (mode)
        {
            case "text":
                Assert.Equal(2, data.GetProperty("totalElements").GetInt32());
                Assert.Contains("CLI view paragraph", data.GetRawText());
                break;
            case "annotated":
                Assert.Equal("annotated", data.GetProperty("view").GetString());
                Assert.Contains("CLI view paragraph", data.GetProperty("content").GetString());
                break;
            case "outline":
                Assert.Equal(1, data.GetProperty("paragraphs").GetInt32());
                Assert.True(data.TryGetProperty("headings", out var headings));
                Assert.Equal(JsonValueKind.Array, headings.ValueKind);
                break;
            case "stats":
                Assert.Equal(1, data.GetProperty("paragraphs").GetInt32());
                Assert.True(data.GetProperty("words").GetInt32() >= 3);
                Assert.True(data.TryGetProperty("styleDistribution", out var styles));
                Assert.Equal(JsonValueKind.Object, styles.ValueKind);
                break;
            case "issues":
                Assert.True(data.TryGetProperty("count", out _));
                Assert.Equal(JsonValueKind.Array, data.GetProperty("issues").ValueKind);
                break;
        }
    }

    [Fact]
    public void ViewForms_ReportsSdtFieldsAndUnknownModeUsesInvalidValueEnvelope()
    {
        var path = CreateBlankDocx();

        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "sdt",
            "--prop", "type=text", "--prop", "alias=Name", "--prop", "tag=name",
            "--prop", "text=Alice", "--json"));

        var forms = AssertJsonSuccess(RunOfficeCli("view", path, "forms", "--json"));
        Assert.Equal("none", forms.GetProperty("protection").GetString());
        Assert.False(forms.GetProperty("protectionEnforced").GetBoolean());
        var field = Assert.Single(forms.GetProperty("fields").EnumerateArray());
        Assert.Equal("sdt", field.GetProperty("kind").GetString());
        Assert.Equal("/body/sdt[1]", field.GetProperty("path").GetString());
        Assert.Equal("text", field.GetProperty("type").GetString());
        Assert.True(field.GetProperty("editable").GetBoolean());
        Assert.Equal("Name", field.GetProperty("alias").GetString());
        Assert.Equal("Alice", field.GetProperty("value").GetString());

        var unknown = RunOfficeCli("view", path, "unknown-mode", "--json");
        Assert.Equal(1, unknown.ExitCode);
        AssertJsonFailure(unknown, "invalid_value");
        Assert.Contains("forms", unknown.Stdout);
    }

    [Fact]
    public void CliCommands_CreateAddGetSetQueryRemoveValidate_MinimalClosedLoop()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_cli_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            Assert.True(File.Exists(path));

            var addData = AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=hello cli", "--json"));
            Assert.Contains("Added paragraph", addData.GetString() ?? addData.GetRawText());

            var getData = AssertJsonSuccess(RunOfficeCli("get", path, "/body/p[1]", "--json"));
            var paragraph = AssertSingleResult(getData);
            Assert.Equal("paragraph", paragraph.GetProperty("type").GetString());
            Assert.Equal("hello cli", paragraph.GetProperty("text").GetString());

            AssertJsonSuccess(RunOfficeCli("set", path, "/body/p[1]", "--prop", "align=center", "--json"));
            var formattedData = AssertJsonSuccess(RunOfficeCli("get", path, "/body/p[1]", "--json"));
            var formatted = AssertSingleResult(formattedData);
            Assert.Equal("center", formatted.GetProperty("format").GetProperty("align").GetString());

            var queryData = AssertJsonSuccess(RunOfficeCli(
                "query", path, "paragraph", "--find", "HELLO", "--json"));
            var queryResult = AssertSingleResult(queryData);
            Assert.True(queryResult.GetProperty("childCount").GetInt32() > 0);
            Assert.Equal(JsonValueKind.Array, queryResult.GetProperty("children").ValueKind);

            AssertJsonSuccess(RunOfficeCli("remove", path, "/body/p[1]", "--json"));
            var emptyQuery = AssertJsonSuccess(RunOfficeCli(
                "query", path, "paragraph", "--find", "hello", "--json"));
            Assert.Equal(0, emptyQuery.GetProperty("matches").GetInt32());

            AssertJsonSuccess(RunOfficeCli("validate", path, "--json"));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void CreateCli_ExistingFileRequiresForceAndForceOverwrites()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_create_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=before force", "--json"));

            var duplicate = RunOfficeCli("create", path, "--json");
            Assert.Equal(1, duplicate.ExitCode);
            AssertJsonFailure(duplicate, "file_exists");

            AssertJsonSuccess(RunOfficeCli("create", path, "--force", "--json"));
            using var handler = new WordHandler(path, editable: false);
            Assert.DoesNotContain("before force", handler.Query("paragraph").Select(p => p.Text));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void CreateCli_OptionMatrix_FixesTypeLocaleAndMinimalBaseline()
    {
        var standardPath = Path.Combine(Path.GetTempPath(), $"officecli_create_standard_{Guid.NewGuid():N}.docx");
        var inferredBasePath = Path.Combine(Path.GetTempPath(), $"officecli_create_inferred_{Guid.NewGuid():N}");
        var inferredPath = inferredBasePath + ".docx";
        var cjkPath = Path.Combine(Path.GetTempPath(), $"officecli_create_cjk_{Guid.NewGuid():N}.docx");
        var rtlPath = Path.Combine(Path.GetTempPath(), $"officecli_create_rtl_{Guid.NewGuid():N}.docx");
        var minimalPath = Path.Combine(Path.GetTempPath(), $"officecli_create_minimal_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", standardPath, "--json"));
            AssertJsonSuccess(RunOfficeCli("create", inferredBasePath, "--type", "docx", "--json"));
            AssertJsonSuccess(RunOfficeCli("create", cjkPath, "--locale", "zh-CN", "--json"));
            AssertJsonSuccess(RunOfficeCli("create", rtlPath, "--locale", "ar-SA", "--json"));
            AssertJsonSuccess(RunOfficeCli("create", minimalPath, "--minimal", "--json"));

            Assert.True(File.Exists(inferredPath));
            foreach (var path in new[] { standardPath, inferredPath, cjkPath, rtlPath, minimalPath })
                AssertJsonSuccess(RunOfficeCli("validate", path, "--json"));

            using (var standard = WordprocessingDocument.Open(standardPath, false))
            {
                var main = standard.MainDocumentPart!;
                Assert.NotNull(main.ThemePart);
                var styles = main.StyleDefinitionsPart!.Styles!;
                Assert.Contains(styles.Elements<Style>(), style => style.StyleId?.Value == "Normal");
                Assert.False(string.IsNullOrWhiteSpace(styles.DocDefaults!.RunPropertiesDefault!
                    .RunPropertiesBaseStyle!.GetFirstChild<RunFonts>()!.Ascii!.Value));
            }

            using (var cjk = WordprocessingDocument.Open(cjkPath, false))
            {
                var main = cjk.MainDocumentPart!;
                var settings = main.DocumentSettingsPart!.Settings!;
                Assert.Equal("zh-CN", settings.GetFirstChild<ThemeFontLanguages>()!.EastAsia!.Value);
                var fonts = main.StyleDefinitionsPart!.Styles!.DocDefaults!.RunPropertiesDefault!
                    .RunPropertiesBaseStyle!.GetFirstChild<RunFonts>();
                Assert.False(string.IsNullOrWhiteSpace(fonts?.EastAsia?.Value));
            }

            using (var rtl = WordprocessingDocument.Open(rtlPath, false))
            {
                var main = rtl.MainDocumentPart!;
                var settings = main.DocumentSettingsPart!.Settings!;
                Assert.Equal("ar-SA", settings.GetFirstChild<ThemeFontLanguages>()!.Bidi!.Value);
                Assert.NotNull(main.Document!.Body!.GetFirstChild<SectionProperties>()!.GetFirstChild<BiDi>());
            }

            using (var minimal = WordprocessingDocument.Open(minimalPath, false))
            {
                var main = minimal.MainDocumentPart!;
                Assert.Null(main.ThemePart);
                var styles = main.StyleDefinitionsPart!.Styles!;
                Assert.DoesNotContain(styles.Elements<Style>(), style => style.StyleId?.Value == "Normal");
                Assert.Equal("Times New Roman", styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!
                    .GetFirstChild<RunFonts>()!.Ascii!.Value);
            }
        }
        finally
        {
            foreach (var path in new[] { standardPath, inferredBasePath, inferredPath, cjkPath, rtlPath, minimalPath })
            {
                try { File.Delete(path); } catch { }
            }
        }
    }

    [Fact]
    public void CliWordFlow_HandlesUnicodeContentAndPathWithWhitespace()
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"\u6587\u6863 {Guid.NewGuid():N} \u6d4b\u8bd5.docx");
        var text = "\u4e2d\u6587 \u6587\u6863 / \u05e9\u05dc\u05d5\u05dd / \u0645\u0631\u062d\u0628\u0627";

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--locale", "zh-CN", "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", $"--prop", $"text={text}", "--json"));

            var query = AssertJsonSuccess(RunOfficeCli("query", path, "paragraph", "--json"));
            Assert.Contains(text, query.GetRawText());

            var view = AssertJsonSuccess(RunOfficeCli("view", path, "text", "--json"));
            Assert.Contains(text, view.GetRawText());
            AssertJsonSuccess(RunOfficeCli("validate", path, "--json"));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void GetCli_DepthAndMissingPathKeepStableJsonShape()
    {
        var path = CreateBlankDocx();

        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=depth target", "--json"));

        var depthZero = AssertSingleResult(AssertJsonSuccess(
            RunOfficeCli("get", path, "/body", "--depth", "0", "--json")));
        Assert.Equal("body", depthZero.GetProperty("type").GetString());
        Assert.Empty(depthZero.GetProperty("children").EnumerateArray());

        var depthOne = AssertSingleResult(AssertJsonSuccess(
            RunOfficeCli("get", path, "/body", "--depth", "1", "--json")));
        Assert.Contains(depthOne.GetProperty("children").EnumerateArray(),
            child => child.GetProperty("type").GetString() == "paragraph");

        var capped = RunOfficeCli("get", path, "/body", "--depth", "999999", "--json");
        Assert.Equal(0, capped.ExitCode);
        AssertJsonSuccess(capped);

        var missing = RunOfficeCli("get", path, "/body/p[99]", "--json");
        Assert.Equal(1, missing.ExitCode);
        using var missingDocument = JsonDocument.Parse(missing.Stdout);
        var missingRoot = missingDocument.RootElement;
        Assert.False(missingRoot.GetProperty("success").GetBoolean());
        Assert.True(missingRoot.TryGetProperty("error", out _) || missingRoot.TryGetProperty("message", out _),
            missing.Stdout);
        Assert.Contains("/body/p[99]", missing.Stdout);
    }

    [Fact]
    public void ViewCli_SelectionAndIssueOptionsRespectCurrentLimits()
    {
        var path = CreateBlankDocx();
        foreach (var text in new[] { "view first", "view second", "view third", "view fourth" })
        {
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", $"text={text}", "--json"));
        }

        var window = AssertJsonSuccess(RunOfficeCli(
            "view", path, "text", "--start", "2", "--end", "3", "--json"));
        Assert.Equal(5, window.GetProperty("totalElements").GetInt32());
        var windowElements = window.GetProperty("elements");
        Assert.Equal(2, windowElements.GetArrayLength());
        Assert.Contains("view second", windowElements.GetRawText());
        Assert.Contains("view third", windowElements.GetRawText());
        Assert.DoesNotContain("view first", windowElements.GetRawText());
        Assert.DoesNotContain("view fourth", windowElements.GetRawText());

        var limited = AssertJsonSuccess(RunOfficeCli(
            "view", path, "text", "--start", "1", "--max-lines", "2", "--json"));
        Assert.Equal(2, limited.GetProperty("elements").GetArrayLength());

        var annotated = AssertJsonSuccess(RunOfficeCli(
            "view", path, "annotated", "--start", "2", "--end", "2", "--json"));
        var annotatedContent = annotated.GetProperty("content").GetString()!;
        Assert.Contains("view second", annotatedContent);
        Assert.DoesNotContain("view first", annotatedContent);

        var noIssues = AssertJsonSuccess(RunOfficeCli(
            "view", path, "issues", "--type", "structure", "--limit", "0", "--json"));
        Assert.Equal(0, noIssues.GetProperty("count").GetInt32());
        Assert.Empty(noIssues.GetProperty("issues").EnumerateArray());
    }

    [Fact]
    public void ViewHtmlCli_PageAndOutKeepEquivalentPreviewContent()
    {
        var path = CreateBlankDocx();
        var outputPath = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Guid.NewGuid():N}.html");
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=HTML \u8F93\u51FA", "--json"));

        try
        {
            var stdoutPreview = RunOfficeCli("view", path, "html", "--page", "1");
            Assert.True(stdoutPreview.ExitCode == 0, stdoutPreview.Stderr);
            Assert.Contains("data-path=\"/body/p[1]\"", stdoutPreview.Stdout);
            Assert.Contains("HTML", stdoutPreview.Stdout);

            var fileResult = RunOfficeCli("view", path, "html", "--page", "1", "--out", outputPath);
            Assert.True(fileResult.ExitCode == 0, fileResult.Stderr);
            Assert.True(File.Exists(outputPath));
            var filePreview = File.ReadAllText(outputPath);
            Assert.Contains("data-path=\"/body/p[1]\"", filePreview);
            Assert.Contains("HTML", filePreview);
            Assert.Equal(stdoutPreview.Stdout, filePreview);
        }
        finally
        {
            try { File.Delete(outputPath); } catch { }
        }
    }

    [Fact]
    public void ViewCli_ScreenshotProducesArtifactOrExplainsBackendUnavailable()
    {
        var path = CreateBlankDocx();
        var screenshotPath = Path.Combine(Path.GetTempPath(), $"officecli_screenshot_{Guid.NewGuid():N}.png");
        var gridPath = Path.Combine(Path.GetTempPath(), $"officecli_screenshot_grid_{Guid.NewGuid():N}.png");
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=visual output", "--json"));

        try
        {
            var screenshot = RunOfficeCli(
                "view", path, "screenshot", "--render", "html", "--page", "1",
                "--screenshot-width", "640", "--screenshot-height", "480", "--out", screenshotPath);
            if (screenshot.ExitCode == 0)
            {
                Assert.True(File.Exists(screenshotPath));
                var png = File.ReadAllBytes(screenshotPath);
                Assert.True(png.Length > 8);
                Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, png[..8]);
            }
            else
            {
                Assert.False(string.IsNullOrWhiteSpace(screenshot.Stdout + screenshot.Stderr));
                Assert.Contains("browser", (screenshot.Stdout + screenshot.Stderr).ToLowerInvariant());
            }

            var grid = RunOfficeCli(
                "view", path, "screenshot", "--render", "html", "--grid", "2",
                "--screenshot-width", "640", "--out", gridPath);
            if (grid.ExitCode == 0)
            {
                Assert.True(File.Exists(gridPath));
                Assert.True(new FileInfo(gridPath).Length > 8);
            }
            else
            {
                Assert.False(string.IsNullOrWhiteSpace(grid.Stdout + grid.Stderr));
                Assert.Contains("browser", (grid.Stdout + grid.Stderr).ToLowerInvariant());
            }

            if (!OperatingSystem.IsWindows())
            {
                var native = RunOfficeCli("view", path, "screenshot", "--render", "native");
                Assert.NotEqual(0, native.ExitCode);
                Assert.Contains("Windows", native.Stdout + native.Stderr);
            }

        }
        finally
        {
            try { File.Delete(screenshotPath); } catch { }
            try { File.Delete(gridPath); } catch { }
        }
    }

    [Fact]
    public void BatchCli_InlineFileAndStdinInputsProduceEquivalentWordResults()
    {
        const string items = "[{\"command\":\"add\",\"parent\":\"/body\",\"type\":\"paragraph\",\"props\":{\"text\":\"batch source\"}},{\"command\":\"query\",\"selector\":\"paragraph\"}]";
        var inputPath = Path.Combine(Path.GetTempPath(), $"officecli_batch_{Guid.NewGuid():N}.json");
        var inlinePath = CreateBlankDocx();
        var filePath = CreateBlankDocx();
        var stdinPath = CreateBlankDocx();
        var explicitStdinPath = CreateBlankDocx();

        try
        {
            File.WriteAllText(inputPath, items, Encoding.UTF8);

            AssertBatchSourceSuccess(RunOfficeCli("batch", inlinePath, "--commands", items, "--json"));
            AssertBatchSourceSuccess(RunOfficeCli("batch", filePath, "--input", inputPath, "--json"));
            AssertBatchSourceSuccess(RunOfficeCliWithInput(items, "batch", stdinPath, "--json"));
            AssertBatchSourceSuccess(RunOfficeCliWithInput(items, "batch", explicitStdinPath, "--input", "-", "--json"));

            foreach (var path in new[] { inlinePath, filePath, stdinPath, explicitStdinPath })
            {
                using var handler = new WordHandler(path, editable: false);
                Assert.Contains("batch source", handler.Query("paragraph").Select(node => node.Text));
            }
        }
        finally
        {
            try { File.Delete(inputPath); } catch { }
        }
    }

    [Fact]
    public void BatchCli_InvalidInputFailsBeforeMutatingTarget()
    {
        var path = CreateBlankDocx();
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=before batch error", "--json"));
        var before = File.ReadAllBytes(path);

        var objectInput = RunOfficeCli("batch", path, "--commands", "{\"command\":\"get\",\"path\":\"/\"}", "--json");
        Assert.Equal(1, objectInput.ExitCode);
        AssertJsonFailure(objectInput, "invalid_input");
        Assert.Equal(before, File.ReadAllBytes(path));

        var unknownField = RunOfficeCli("batch", path, "--commands",
            "[{\"command\":\"add\",\"parent\":\"/body\",\"type\":\"paragraph\",\"unknown\":\"x\"}]", "--json");
        Assert.Equal(1, unknownField.ExitCode);
        AssertJsonFailure(unknownField, "internal_error");
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Fact]
    public void BatchCli_EmptyArrayAndUtf8BomInputKeepEnvelopeAndReplaySemantics()
    {
        var emptyPath = CreateBlankDocx();
        var bomPath = CreateBlankDocx();
        var inputPath = Path.Combine(Path.GetTempPath(), $"officecli_batch_bom_{Guid.NewGuid():N}.json");
        const string bomItems = "[{\"command\":\"add\",\"parent\":\"/body\",\"type\":\"paragraph\",\"props\":{\"text\":\"BOM batch\"}}]";

        try
        {
            var emptyData = AssertJsonSuccess(RunOfficeCli(
                "batch", emptyPath, "--commands", "[]", "--json"));
            Assert.Equal(0, emptyData.GetProperty("summary").GetProperty("total").GetInt32());
            Assert.Equal(0, emptyData.GetProperty("results").GetArrayLength());

            var utf8WithBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
            File.WriteAllText(inputPath, bomItems, utf8WithBom);
            AssertBatchSourceSuccess(RunOfficeCli("batch", bomPath, "--input", inputPath, "--json"), expectedTotal: 1);
            using var handler = new WordHandler(bomPath, editable: false);
            Assert.Contains("BOM batch", handler.Query("paragraph").Select(node => node.Text));
        }
        finally
        {
            try { File.Delete(inputPath); } catch { }
        }
    }

    [Fact]
    public void BatchCli_RejectsJsonlAndConflictingSourcesBeforeMutatingTarget()
    {
        var jsonlPath = Path.Combine(Path.GetTempPath(), $"officecli_batch_conflict_{Guid.NewGuid():N}.json");
        var path = CreateBlankDocx();
        var before = File.ReadAllBytes(path);
        const string jsonl = "{\"command\":\"add\",\"parent\":\"/body\",\"type\":\"paragraph\",\"props\":{\"text\":\"jsonl one\"}}\n"
                          + "{\"command\":\"add\",\"parent\":\"/body\",\"type\":\"paragraph\",\"props\":{\"text\":\"jsonl two\"}}\n";

        try
        {
            var jsonlResult = RunOfficeCliWithInput(jsonl, "batch", path, "--json");
            Assert.Equal(1, jsonlResult.ExitCode);
            AssertJsonFailure(jsonlResult, "invalid_json");
            Assert.Equal(before, File.ReadAllBytes(path));

            File.WriteAllText(jsonlPath, "[]", Encoding.UTF8);
            var conflictingSources = RunOfficeCli(
                "batch", path, "--commands", "[]", "--input", jsonlPath, "--json");
            Assert.Equal(1, conflictingSources.ExitCode);
            AssertJsonFailure(conflictingSources, "internal_error");
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            try { File.Delete(jsonlPath); } catch { }
        }
    }

    [Fact]
    public void ResidentLock_RejectsForceCreateAndRecoversAfterUnexpectedExit()
    {
        var path = CreateBlankDocx();
        var startInfo = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        startInfo.ArgumentList.Add(FindOfficeCliDll());
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "off";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        _ = resident.StandardOutput.ReadToEndAsync();
        _ = resident.StandardError.ReadToEndAsync();
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");

            var locked = RunOfficeCli("create", path, "--force", "--json");
            Assert.Equal(1, locked.ExitCode);
            AssertJsonFailure(locked, "file_locked");

            resident.Kill(entireProcessTree: true);
            Assert.True(resident.WaitForExit(5_000), "resident did not exit after kill");
            Assert.True(
                SpinWait.SpinUntil(() => !OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident pipe remained reachable after process exit");

            AssertJsonSuccess(RunOfficeCli("create", path, "--force", "--json"));
            Assert.True(File.Exists(path));
        }
        finally
        {
            if (!resident.HasExited)
            {
                try { OfficeCli.ResidentClient.SendClose(path); } catch { }
                if (!resident.WaitForExit(5_000))
                {
                    try { resident.Kill(entireProcessTree: true); } catch { }
                }
            }
        }
    }

    [Fact]
    public void AddCli_PositionOptionsAndFromConflictFollowCurrentRules()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_add_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--prop", "text=A", "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--prop", "text=C", "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--index", "1", "--prop", "text=B", "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--before", "/body/p[1]", "--prop", "text=X", "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--after", "/body/p[4]", "--prop", "text=D", "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--from", "/body/p[2]", "--json"));

            using (var handler = new WordHandler(path, editable: false))
            {
                Assert.Equal(["X", "A", "B", "C", "D", "A"], handler.Query("paragraph").Select(p => p.Text ?? "").ToArray());
            }

            var fromWithProp = RunOfficeCli(
                "add", path, "/body", "--from", "/body/p[2]", "--prop", "text=ignored", "--json");
            Assert.Equal(1, fromWithProp.ExitCode);
            AssertJsonFailure(fromWithProp, "invalid_argument");

            var conflictingPosition = RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--index", "0", "--before", "/body/p[1]", "--json");
            Assert.Equal(1, conflictingPosition.ExitCode);
            AssertJsonFailure(conflictingPosition, "invalid_argument");
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void AddCli_RepresentativeWordTypesReturnQueryablePathsAndValidPackage()
    {
        var path = CreateBlankDocx();

        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=CLI type host", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body/p[1]", "--type", "run", "--prop", "text=inline run", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "table", "--prop", "data=A,B", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "picture", "--prop", $"src={TinyPngDataUri}",
            "--prop", "alt=CLI picture", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "sdt", "--prop", "type=text", "--prop", "text=CLI control", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body/p[1]", "--type", "field", "--prop", "fieldType=page", "--prop", "text=1", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body/p[1]", "--type", "hyperlink", "--prop", "url=https://example.com",
            "--prop", "text=CLI link", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body/p[1]", "--type", "bookmark", "--prop", "name=CliMark", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/", "--type", "header", "--prop", "text=CLI header", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/", "--type", "footer", "--prop", "text=CLI footer", "--json"));

        foreach (var selector in new[]
        {
            "paragraph", "run", "table", "picture", "sdt", "field", "hyperlink", "bookmark", "header", "footer"
        })
        {
            var data = AssertJsonSuccess(RunOfficeCli("query", path, selector, "--json"));
            Assert.True(data.GetProperty("matches").GetInt32() > 0, selector);
        }

        AssertJsonSuccess(RunOfficeCli("validate", path, "--json"));
    }

    [Fact]
    public void AddCli_BarePropertyWarnsAndDoesNotApplyText()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_bare_prop_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));

            var result = RunOfficeCli("add", path, "/body", "--type", "paragraph", "text=bare", "--json");
            Assert.Equal(2, result.ExitCode);
            var data = AssertJsonSuccess(result, expectedExitCode: 2);
            Assert.Contains("Added paragraph", data.GetString() ?? data.GetRawText());
            Assert.Contains("WARNING: Properties specified without --prop flag.", result.Stderr);

            using var handler = new WordHandler(path, editable: false);
            Assert.DoesNotContain("bare", handler.Query("paragraph").Select(p => p.Text));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void SetCli_ReportsMissingBareAndUnsupportedProps()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_set_props_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=set target", "--json"));

            var missing = RunOfficeCli("set", path, "/body/p[1]", "--json");
            Assert.Equal(1, missing.ExitCode);
            AssertJsonWarningFailure(missing, "missing_property");

            var bare = RunOfficeCli("set", path, "/body/p[1]", "align=right", "--json");
            Assert.Equal(2, bare.ExitCode);
            AssertJsonWarningFailure(bare, "missing_prop_flag");

            var unsupported = RunOfficeCli(
                "set", path, "/body/p[1]", "--prop", "align=right", "--prop", "nosuch=1", "--json");
            Assert.Equal(2, unsupported.ExitCode);
            var data = AssertJsonSuccess(unsupported, expectedExitCode: 2);
            Assert.Contains("Updated /body/p[1]: align=right", data.GetString() ?? data.GetRawText());
            AssertJsonWarning(unsupported, "unsupported_property");

            var uppercase = AssertJsonSuccess(RunOfficeCli(
                "set", path, "/body/p[1]", "--prop", "ALIGN=left", "--json"));
            Assert.Contains("ALIGN=left", uppercase.GetString() ?? uppercase.GetRawText());

            var legacyAlias = AssertJsonSuccess(RunOfficeCli(
                "set", path, "/body/p[1]", "--prop", "alignment=justify", "--json"));
            Assert.Contains("alignment=justify", legacyAlias.GetString() ?? legacyAlias.GetRawText());

            using var handler = new WordHandler(path, editable: false);
            Assert.Equal("justify", handler.Get("/body/p[1]").Format["align"]);
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void SetCli_ScopedSelectorUpdatesEveryMatchingParagraph()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_set_selector_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=first target", "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=second target", "--json"));

            var set = AssertJsonSuccess(RunOfficeCli(
                "set", path, "/body/p[text~=target]", "--prop", "align=center", "--json"));
            Assert.Contains("Updated /body/p[text~=target]: align=center", set.GetString() ?? set.GetRawText());

            var first = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
                "get", path, "/body/p[1]", "--json")));
            var second = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
                "get", path, "/body/p[2]", "--json")));
            Assert.Equal("center", first.GetProperty("format").GetProperty("align").GetString());
            Assert.Equal("center", second.GetProperty("format").GetProperty("align").GetString());
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void QueryCli_ComplexSelectorsReturnHydratedJsonResults()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_query_complex_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=plain", "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=bold", "--prop", "bold=true", "--json"));

            var childData = AssertJsonSuccess(RunOfficeCli(
                "query", path, "paragraph > run[bold=true]", "--json"));
            var childRun = AssertSingleResult(childData);
            Assert.Equal("run", childRun.GetProperty("type").GetString());
            Assert.Equal("bold", childRun.GetProperty("text").GetString());
            Assert.True(childRun.GetProperty("format").GetProperty("bold").GetBoolean());

            var orData = AssertJsonSuccess(RunOfficeCli(
                "query", path, "paragraph[text=plain or bold=true]", "--json"));
            Assert.Equal(2, orData.GetProperty("matches").GetInt32());
            var results = orData.GetProperty("results").EnumerateArray().ToArray();
            Assert.Equal(["plain", "bold"], results.Select(r => r.GetProperty("text").GetString() ?? "").ToArray());
            Assert.All(results, result =>
                Assert.Equal(JsonValueKind.Array, result.GetProperty("children").ValueKind));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void RemoveMoveAndCopyCli_FollowCurrentPathAndPositionRules()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_move_copy_{Guid.NewGuid():N}.docx");
        string[] BodyTexts()
        {
            using var handler = new WordHandler(path, editable: false);
            return handler.Query("paragraph").Select(p => p.Text ?? "").ToArray();
        }

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            foreach (var text in new[] { "A", "B", "C", "D" })
            {
                AssertJsonSuccess(RunOfficeCli(
                    "add", path, "/body", "--type", "paragraph", "--prop", $"text={text}", "--json"));
            }

            var removeData = AssertJsonSuccess(RunOfficeCli("remove", path, "/body/p", "--json"));
            Assert.Contains("Removed /body/p", removeData.GetString() ?? removeData.GetRawText());
            Assert.Equal(["B", "C", "D"], BodyTexts());

            var moveBeforeData = AssertJsonSuccess(RunOfficeCli(
                "move", path, "/body/p[3]", "--before", "/body/p[1]", "--json"));
            Assert.Contains("Moved to /body/p[1]", moveBeforeData.GetString() ?? moveBeforeData.GetRawText());
            Assert.Equal(["D", "B", "C"], BodyTexts());

            var moveAfterData = AssertJsonSuccess(RunOfficeCli(
                "move", path, "/body/p[1]", "--after", "/body/p[3]", "--json"));
            Assert.Contains("Moved to /body/p[3]", moveAfterData.GetString() ?? moveAfterData.GetRawText());
            Assert.Equal(["B", "C", "D"], BodyTexts());

            var copyData = AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--from", "/body/p[1]", "--index", "2", "--json"));
            Assert.Contains("Copied to /body/p[3]", copyData.GetString() ?? copyData.GetRawText());
            Assert.Equal(["B", "C", "B", "D"], BodyTexts());

            var moveIndexData = AssertJsonSuccess(RunOfficeCli(
                "move", path, "/body/p[4]", "--index", "1", "--json"));
            Assert.Contains("Moved to /body/p[2]", moveIndexData.GetString() ?? moveIndexData.GetRawText());
            Assert.Equal(["B", "D", "C", "B"], BodyTexts());
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void ProtectedDocx_CliGateMatchesCurrentCommandSurface()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_protected_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--prop", "text=locked", "--json"));
            AssertJsonSuccess(RunOfficeCli("set", path, "/", "--prop", "protection=forms", "--json"));

            AssertProtectedFailure(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=blocked add", "--json"));
            AssertProtectedFailure(RunOfficeCli(
                "set", path, "/body/p[1]", "--prop", "align=center", "--json"));

            AssertJsonSuccess(RunOfficeCli(
                "set", path, "/body/p[1]", "--prop", "align=center", "--force", "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=forced add", "--force", "--json"));

            var blockedBatch = RunOfficeCli("batch", path, "--commands", """
                [
                  {"command":"remove","path":"/body/p[1]"},
                  {"command":"raw-set","part":"/document","xpath":"//w:p[1]","action":"setattr","xml":"w:rsidR=00112233"}
                ]
                """, "--json");
            AssertProtectedFailure(blockedBatch);

            AssertJsonSuccess(RunOfficeCli("batch", path, "--force", "--commands", """
                [
                  {"command":"add","parent":"/body","type":"paragraph","props":{"text":"forced batch"}}
                ]
                """, "--json"));

            AssertJsonSuccess(RunOfficeCli(
                "raw-set", path, "/document", "--xpath", "//w:p[1]", "--action", "setattr", "--xml", "w:rsidR=00112233", "--json"));
            AssertJsonSuccess(RunOfficeCli("remove", path, "/body/p[2]", "--json"));

            using var handler = new WordHandler(path, editable: false);
            var first = handler.Get("/body/p[1]");
            Assert.Equal("center", first.Format["align"]);
            Assert.Contains("forced batch", handler.Query("paragraph").Select(p => p.Text));
            Assert.DoesNotContain("blocked add", handler.Query("paragraph").Select(p => p.Text));
            Assert.Contains("00112233", handler.Raw("/document"));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void GetSave_CliSavesPicturePayloadAndRejectsTextNodePayload()
    {
        var path = CreateBlankDocx();
        string picturePath;
        string olePath;
        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "text node" });
            handler.Add("/body", "picture", null, new()
            {
                ["src"] = TinyPngDataUri,
                ["width"] = "1in",
                ["height"] = "1in",
                ["alt"] = "tiny"
            });
            handler.Add("/body", "ole", null, new()
            {
                ["src"] = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,SGVsbG8=",
                ["progId"] = "Word.Document.12",
                ["name"] = "Embedded Word",
                ["width"] = "2in",
                ["height"] = "1in"
            });
            picturePath = handler.Query("picture").Single().Path!;
            olePath = handler.Query("ole").Single().Path!;
        }

        var pictureQuery = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
            "query", path, "picture", "--json")));
        Assert.Equal("picture", pictureQuery.GetProperty("type").GetString());
        Assert.Equal("inline", pictureQuery.GetProperty("format").GetProperty("wrap").GetString());

        var oleQuery = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
            "query", path, "ole", "--json")));
        Assert.Equal("ole", oleQuery.GetProperty("type").GetString());
        Assert.Equal("Word.Document.12", oleQuery.GetProperty("format").GetProperty("progId").GetString());

        var outDir = Path.Combine(Path.GetTempPath(), $"officecli_save_{Guid.NewGuid():N}");
        Directory.CreateDirectory(outDir);
        try
        {
            var pictureOut = Path.Combine(outDir, "tiny.png");
            var pictureData = AssertJsonSuccess(RunOfficeCli(
                "get", path, picturePath, "--save", pictureOut, "--json"));
            var picture = AssertSingleResult(pictureData);
            var format = picture.GetProperty("format");
            Assert.Equal(pictureOut, format.GetProperty("savedTo").GetString());
            Assert.Equal("image/png", format.GetProperty("savedContentType").GetString());
            Assert.Equal(Convert.FromBase64String(TinyPngDataUri.Split(',')[1]), File.ReadAllBytes(pictureOut));
            Assert.Equal(new FileInfo(pictureOut).Length, format.GetProperty("savedBytes").GetInt64());

            var oleOut = Path.Combine(outDir, "embedded.docx");
            var oleData = AssertJsonSuccess(RunOfficeCli(
                "get", path, olePath, "--save", oleOut, "--json"));
            var ole = AssertSingleResult(oleData);
            var oleFormat = ole.GetProperty("format");
            Assert.Equal(oleOut, oleFormat.GetProperty("savedTo").GetString());
            Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", oleFormat.GetProperty("savedContentType").GetString());
            Assert.Equal(Encoding.UTF8.GetBytes("Hello"), File.ReadAllBytes(oleOut));
            Assert.Equal(new FileInfo(oleOut).Length, oleFormat.GetProperty("savedBytes").GetInt64());

            var textOut = Path.Combine(outDir, "text.bin");
            var textResult = RunOfficeCli("get", path, "/body/p[1]", "--save", textOut, "--json");
            Assert.Equal(1, textResult.ExitCode);
            using var doc = JsonDocument.Parse(textResult.Stdout);
            Assert.False(doc.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("has no binary payload", textResult.Stdout);
            Assert.False(File.Exists(textOut));
        }
        finally
        {
            try { Directory.Delete(outDir, recursive: true); } catch { }
        }
    }

    [Fact]
    public void PictureQueryCli_ReturnsFloatingWrapAndRelativePosition()
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "picture", null, new()
            {
                ["src"] = TinyPngDataUri,
                ["name"] = "floating-cli.png",
                ["wrap"] = "square",
                ["hPosition"] = "2cm",
                ["vPosition"] = "1cm",
                ["hRelative"] = "page",
                ["vRelative"] = "paragraph"
            });
        }

        var picture = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
            "query", path, "picture", "--json")));
        var format = picture.GetProperty("format");
        Assert.Equal("picture", picture.GetProperty("type").GetString());
        Assert.Equal("square", format.GetProperty("wrap").GetString());
        Assert.Equal("2.0cm", format.GetProperty("hPosition").GetString());
        Assert.Equal("1.0cm", format.GetProperty("vPosition").GetString());
        Assert.Equal("page", format.GetProperty("hRelative").GetString());
        Assert.Equal("paragraph", format.GetProperty("vRelative").GetString());
    }

    [Fact]
    public void RemoveCli_OleIndexedScopedShorthandRemovesObjectAndLeavesValidPackage()
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "ole", null, new()
            {
                ["src"] = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,SGVsbG8=",
                ["progId"] = "Word.Document.12",
                ["name"] = "remove through shorthand"
            });
        }

        var removed = AssertJsonSuccess(RunOfficeCli("remove", path, "/body/ole[1]", "--json"));
        Assert.Contains("Removed /body/ole[1]", removed.GetString() ?? removed.GetRawText());

        using var reopened = new WordHandler(path, editable: false);
        Assert.Empty(reopened.Query("ole"));
        Assert.Empty(reopened.Validate());
    }

    [Fact]
    public void SaveCli_NoResidentIsNoOpSuccess()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_save_no_resident_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            var data = AssertJsonSuccess(RunOfficeCli("save", path, "--json"));
            Assert.Contains("already saved", data.GetString() ?? data.GetRawText());
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void ResidentFlushEach_CliMutationIsVisibleOnDiskBeforeSaveOrClose()
    {
        var path = CreateBlankDocx();
        var startInfo = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add(FindOfficeCliDll());
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "each";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        _ = resident.StandardOutput.ReadToEndAsync();
        _ = resident.StandardError.ReadToEndAsync();
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=flush each visible", "--json"));

            using var afterAdd = new WordHandler(path, editable: false);
            Assert.Contains("flush each visible", afterAdd.Query("paragraph").Select(p => p.Text));
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
            if (!resident.WaitForExit(5_000))
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
            }
        }
    }

    [Fact]
    public void ResidentFixedFlushInterval_MutationBecomesVisibleToExternalReader()
    {
        var path = CreateBlankDocx();
        var startInfo = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "1";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=fixed flush visible", "--json"));

            var visible = false;
            var deadline = DateTime.UtcNow.AddSeconds(15);
            while (DateTime.UtcNow < deadline)
            {
                try
                {
                    using var reader = new WordHandler(path, editable: false);
                    if (reader.Query("paragraph").Any(paragraph => paragraph.Text == "fixed flush visible"))
                    {
                        visible = true;
                        break;
                    }
                }
                catch
                {
                    // The resident may hold the file while the background flush is in progress.
                }

                Thread.Sleep(100);
            }
            Assert.True(visible, "fixed-interval resident flush did not reach disk");
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
            if (!resident.WaitForExit(5_000))
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
            }
        }
    }

    [Fact]
    public void ResidentAdaptiveAutoFlush_MutationBecomesVisibleAfterIdleWindow()
    {
        var path = CreateBlankDocx();
        var startInfo = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "auto";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        _ = resident.StandardOutput.ReadToEndAsync();
        _ = resident.StandardError.ReadToEndAsync();
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=adaptive flush visible", "--json"));

            var visible = false;
            var deadline = DateTime.UtcNow.AddSeconds(15);
            while (DateTime.UtcNow < deadline)
            {
                try
                {
                    using var reader = new WordHandler(path, editable: false);
                    if (reader.Query("paragraph").Any(paragraph => paragraph.Text == "adaptive flush visible"))
                    {
                        visible = true;
                        break;
                    }
                }
                catch
                {
                    // The resident may hold the file while the background flush is in progress.
                }

                Thread.Sleep(100);
            }

            Assert.True(visible, "adaptive resident flush did not reach disk");
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
            if (!resident.WaitForExit(5_000))
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
            }
        }
    }

    [Fact]
    public void ResidentBatchFlushEach_MakesBatchMutationVisibleBeforeClose()
    {
        var path = CreateBlankDocx();
        var startInfo = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "each";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");
            AssertJsonSuccess(RunOfficeCli("batch", path, "--commands", """
                [{"command":"add","parent":"/body","type":"paragraph","props":{"text":"batch each visible"}}]
                """, "--json"));

            using var reader = new WordHandler(path, editable: false);
            Assert.Contains("batch each visible", reader.Query("paragraph").Select(paragraph => paragraph.Text));
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
            if (!resident.WaitForExit(5_000))
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
            }
        }
    }

    [Fact]
    public void ResidentScopedSelectorSet_ForwardsThroughCliAndPersistsAfterClose()
    {
        var path = CreateBlankDocx();
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=first resident target", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=second resident target", "--json"));

        var startInfo = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add(FindOfficeCliDll());
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "off";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");

            var set = AssertJsonSuccess(RunOfficeCli(
                "set", path, "/body/p[text~=resident]", "--prop", "align=right", "--json"));
            Assert.Contains("Updated /body/p[text~=resident]: align=right", set.GetString() ?? set.GetRawText());

            var first = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
                "get", path, "/body/p[1]", "--json")));
            var second = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
                "get", path, "/body/p[2]", "--json")));
            Assert.Equal("right", first.GetProperty("format").GetProperty("align").GetString());
            Assert.Equal("right", second.GetProperty("format").GetProperty("align").GetString());

            Assert.True(OfficeCli.ResidentClient.SendClose(path));
            Assert.True(resident.WaitForExit(5_000), "resident did not close");

            using var afterClose = new WordHandler(path, editable: false);
            Assert.Equal("right", afterClose.Get("/body/p[1]").Format["align"]);
            Assert.Equal("right", afterClose.Get("/body/p[2]").Format["align"]);
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
            if (!resident.HasExited)
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
            }
            try { resident.WaitForExit(5_000); } catch { }
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void OpenCli_StartsResidentAndCloseFlushesCliMutation()
    {
        var path = CreateBlankDocx();

        try
        {
            var opened = AssertJsonSuccess(RunOfficeCliAppHost("open", path, "--json"));
            Assert.Contains("resident", opened.GetString() ?? opened.GetRawText(), StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficeCli.ResidentClient.TryConnect(path, out _));

            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=opened through CLI", "--json"));
            var live = AssertJsonSuccess(RunOfficeCli("query", path, "paragraph", "--find", "opened through CLI", "--json"));
            Assert.Equal(1, live.GetProperty("matches").GetInt32());

            AssertJsonSuccess(RunOfficeCli("close", path, "--json"));
            SpinWait.SpinUntil(() => !OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5));

            using var reopened = new WordHandler(path, editable: false);
            Assert.Contains("opened through CLI", reopened.Query("paragraph").Select(p => p.Text));
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
        }
    }

    [Fact]
    public void SwapCli_ResidentForwardPersistsReorderedParagraphs()
    {
        var path = CreateBlankDocx();
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=resident A", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=resident B", "--json"));

        var startInfo = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "off";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");

            var swapped = AssertJsonSuccess(RunOfficeCli(
                "swap", path, "/body/p[1]", "/body/p[2]", "--json"));
            Assert.Contains("Swapped", swapped.GetString() ?? swapped.GetRawText());

            var live = AssertJsonSuccess(RunOfficeCli("query", path, "paragraph", "--json"));
            Assert.Equal(["resident B", "resident A"], live.GetProperty("results").EnumerateArray()
                .Select(result => result.GetProperty("text").GetString() ?? "").ToArray());

            Assert.True(OfficeCli.ResidentClient.SendClose(path));
            Assert.True(resident.WaitForExit(5_000), "resident did not close");

            using var reopened = new WordHandler(path, editable: false);
            Assert.Equal(["resident B", "resident A"], reopened.Query("paragraph")
                .Select(paragraph => paragraph.Text ?? "").ToArray());
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
            if (!resident.HasExited)
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
            }
            try { resident.WaitForExit(5_000); } catch { }
        }
    }

    [Fact]
    public void MoveAndCopyCli_ResidentForwardChangesMemoryAndExplicitSavePersistsIt()
    {
        var path = CreateBlankDocx();
        foreach (var text in new[] { "resident move A", "resident move B", "resident move C" })
        {
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", $"text={text}", "--json"));
        }

        var startInfo = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "off";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");

            var moved = AssertJsonSuccess(RunOfficeCli(
                "move", path, "/body/p[3]", "--before", "/body/p[1]", "--json"));
            Assert.Contains("Moved to /body/p[1]", moved.GetString() ?? moved.GetRawText());
            Assert.True(OfficeCli.ResidentClient.TryConnect(path, out _), "move was not forwarded to the live resident");

            var copied = AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--from", "/body/p[1]", "--index", "2", "--json"));
            Assert.Contains("Copied to /body/p[3]", copied.GetString() ?? copied.GetRawText());
            Assert.True(OfficeCli.ResidentClient.TryConnect(path, out _), "copy was not forwarded to the live resident");

            var live = AssertJsonSuccess(RunOfficeCli("query", path, "paragraph", "--json"));
            var liveTexts = live.GetProperty("results").EnumerateArray()
                .Select(result => result.GetProperty("text").GetString() ?? "")
                .ToArray();
            Assert.True(
                liveTexts.SequenceEqual(["resident move C", "resident move A", "resident move C", "resident move B"]),
                live.GetRawText());

            var saved = RunOfficeCli("save", path, "--json");
            var savedData = AssertJsonSuccess(saved);
            Assert.Contains("Saved", savedData.GetString() ?? savedData.GetRawText());

            Assert.True(OfficeCli.ResidentClient.SendClose(path));
            Assert.True(resident.WaitForExit(5_000), "resident did not close");

            using var reopened = new WordHandler(path, editable: false);
            Assert.Equal(
                ["resident move C", "resident move A", "resident move C", "resident move B"],
                reopened.Query("paragraph").Select(paragraph => paragraph.Text ?? "").ToArray());
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
            if (!resident.HasExited)
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
            }
            try { resident.WaitForExit(5_000); } catch { }
        }
    }

    [Fact]
    public void RemoveCli_ResidentForwardPersistsDeletion()
    {
        var path = CreateBlankDocx();
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=resident remove", "--json"));
        AssertJsonSuccess(RunOfficeCli(
            "add", path, "/body", "--type", "paragraph", "--prop", "text=resident keep", "--json"));

        var startInfo = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("__resident-serve__");
        startInfo.ArgumentList.Add(path);
        startInfo.Environment["OFFICECLI_RESIDENT_FLUSH"] = "off";
        startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = "60";

        using var resident = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start resident");
        try
        {
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.ResidentClient.TryConnect(path, out _), TimeSpan.FromSeconds(5)),
                "resident did not start");

            var removed = AssertJsonSuccess(RunOfficeCli(
                "remove", path, "/body/p[1]", "--json"));
            Assert.Contains("Removed /body/p[1]", removed.GetString() ?? removed.GetRawText());

            var live = AssertJsonSuccess(RunOfficeCli("query", path, "paragraph", "--json"));
            Assert.Equal(["resident keep"], live.GetProperty("results").EnumerateArray()
                .Select(result => result.GetProperty("text").GetString() ?? "").ToArray());

            Assert.True(OfficeCli.ResidentClient.SendClose(path));
            Assert.True(resident.WaitForExit(5_000), "resident did not close");

            using var reopened = new WordHandler(path, editable: false);
            Assert.Equal(["resident keep"], reopened.Query("paragraph")
                .Select(paragraph => paragraph.Text ?? "").ToArray());
        }
        finally
        {
            try { OfficeCli.ResidentClient.SendClose(path); } catch { }
            if (!resident.HasExited)
            {
                try { resident.Kill(entireProcessTree: true); } catch { }
            }
            try { resident.WaitForExit(5_000); } catch { }
        }
    }

    [Fact]
    public void DumpCli_OutFileWritesBareBatchArrayAndStdoutReportsMetadata()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_dump_{Guid.NewGuid():N}.docx");
        var outPath = Path.Combine(Path.GetTempPath(), $"officecli_dump_{Guid.NewGuid():N}.json");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=dump target", "--json"));

            var data = AssertJsonSuccess(RunOfficeCli("dump", path, "/body", "--out", outPath, "--json"));
            Assert.Equal(outPath, data.GetProperty("outputFile").GetString());
            Assert.True(data.GetProperty("itemCount").GetInt32() > 0);

            using var dumped = JsonDocument.Parse(File.ReadAllText(outPath));
            Assert.Equal(JsonValueKind.Array, dumped.RootElement.ValueKind);
            Assert.Contains("dump target", dumped.RootElement.GetRawText());

            var stdoutDump = AssertJsonSuccess(RunOfficeCli("dump", path, "/body", "--out", "-", "--json"));
            Assert.Equal(JsonValueKind.Array, stdoutDump.ValueKind);
            Assert.Contains("dump target", stdoutDump.GetRawText());

            var invalid = RunOfficeCli("dump", path, "--format", "xml", "--json");
            Assert.Equal(1, invalid.ExitCode);
            AssertJsonFailure(invalid, "invalid_format");
        }
        finally
        {
            try { File.Delete(path); } catch { }
            try { File.Delete(outPath); } catch { }
        }
    }

    [Fact]
    public void DumpCli_JsonWarningsStayInEnvelopeAndStderrWithoutPollutingBatchData()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "orphan note host" });
            handler.Add(paragraphPath, "footnote", null, new() { ["text"] = "orphan footnote" });
            handler.RawSet("/document", "//w:footnoteReference", "remove", null);
        }

        var result = RunOfficeCli("dump", path, "/", "--json");
        Assert.Equal(0, result.ExitCode);
        using var document = JsonDocument.Parse(result.Stdout);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), result.Stdout);
        Assert.Equal(JsonValueKind.Array, root.GetProperty("data").ValueKind);
        Assert.Contains(root.GetProperty("warnings").EnumerateArray(), warning =>
            warning.GetProperty("code").GetString() == "unsupported_element"
            && warning.GetProperty("message").GetString()!.Contains("footnote", StringComparison.OrdinalIgnoreCase));
        Assert.Contains("warning:", result.Stderr, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("warning:", root.GetProperty("data").GetRawText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RawCli_RawSetAndAddPartExposeCurrentProtocol()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_raw_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=raw target", "--json"));

            var raw = AssertJsonSuccess(RunOfficeCli("raw", path, "/document", "--json"));
            Assert.Contains("<w:document", raw.GetString() ?? raw.GetRawText());
            Assert.Contains("raw target", raw.GetString() ?? raw.GetRawText());

            var zipUri = AssertJsonSuccess(RunOfficeCli(
                "raw", path, "/word/document.xml?view=1#root", "--json"));
            Assert.Contains("<w:document", zipUri.GetString() ?? zipUri.GetRawText());

            var missingPart = RunOfficeCli("raw", path, "/word/missing.xml", "--json");
            Assert.Equal(1, missingPart.ExitCode);
            AssertJsonFailure(missingPart, "invalid_value");

            var rawSet = AssertJsonSuccess(RunOfficeCli(
                "raw-set", path, "/document", "--xpath", "//w:p[1]", "--action", "setattr", "--xml", "w:rsidR=00112233", "--json"));
            Assert.Contains("raw-set applied", rawSet.GetString() ?? rawSet.GetRawText());

            var addPart = AssertJsonSuccess(RunOfficeCli("add-part", path, "/", "--type", "chart", "--json"));
            var addPartText = addPart.GetString() ?? addPart.GetRawText();
            Assert.Contains("relId=", addPartText);
            Assert.Contains("path=/chart[1]", addPartText);

            var addPartAgain = AssertJsonSuccess(RunOfficeCli("add-part", path, "/", "--type", "chart", "--json"));
            var addPartAgainText = addPartAgain.GetString() ?? addPartAgain.GetRawText();
            Assert.Contains("relId=", addPartAgainText);
            Assert.Contains("path=/chart[2]", addPartAgainText);
            Assert.NotEqual(addPartText, addPartAgainText);

            var chartRaw = AssertJsonSuccess(RunOfficeCli("raw", path, "/chart[1]", "--json"));
            Assert.Contains("<c:chartSpace", chartRaw.GetString() ?? chartRaw.GetRawText());
            var chartRawAgain = AssertJsonSuccess(RunOfficeCli("raw", path, "/chart[2]", "--json"));
            Assert.Contains("<c:chartSpace", chartRawAgain.GetString() ?? chartRawAgain.GetRawText());

            var validation = RunOfficeCli("validate", path, "--json");
            Assert.Equal(1, validation.ExitCode);
            using (var validationDocument = JsonDocument.Parse(validation.Stdout))
            {
                var validationRoot = validationDocument.RootElement;
                Assert.False(validationRoot.GetProperty("success").GetBoolean());
                var errors = validationRoot.GetProperty("data").GetProperty("errors");
                Assert.Equal(2, errors.GetArrayLength());
                Assert.Equal(["/word/charts/chart1.xml", "/word/charts/chart2.xml"], errors.EnumerateArray()
                    .Select(error => error.GetProperty("part").GetString() ?? string.Empty).ToArray());
                Assert.All(errors.EnumerateArray(), error =>
                {
                    Assert.Equal("Schema", error.GetProperty("type").GetString());
                    Assert.Contains("incomplete content", error.GetProperty("description").GetString(),
                        StringComparison.OrdinalIgnoreCase);
                });
            }

            var invalidRawSet = RunOfficeCli(
                "raw-set", path, "/document", "--xpath", "//w:p[99]", "--action", "setattr", "--xml", "w:rsidR=44556677", "--json");
            Assert.Equal(1, invalidRawSet.ExitCode);
            Assert.Contains("XPath matched no elements", invalidRawSet.Stdout);

            var invalidXml = RunOfficeCli(
                "raw-set", path, "/document", "--xpath", "//w:p[1]", "--action", "replace", "--xml", "<w:p>", "--json");
            Assert.Equal(1, invalidXml.ExitCode);
            AssertJsonFailure(invalidXml, "internal_error");
            var afterInvalidXml = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
                "get", path, "/body/p[1]", "--json")));
            Assert.Equal("raw target", afterInvalidXml.GetProperty("text").GetString());

            using var handler = new WordHandler(path, editable: false);
            Assert.Contains("00112233", handler.Raw("/document"));
            Assert.DoesNotContain("44556677", handler.Raw("/document"));
            Assert.Contains(handler.Validate(), error =>
                error.Part == "/word/charts/chart1.xml"
                && error.Description.Contains("incomplete content", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void ErrorCli_FilePathWatchAndSelectorErrorsExposeCurrentJsonShape()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_errors_{Guid.NewGuid():N}.docx");
        var missingPath = Path.Combine(Path.GetTempPath(), $"officecli_missing_{Guid.NewGuid():N}.docx");
        var unsupportedPath = Path.Combine(Path.GetTempPath(), $"officecli_unsupported_{Guid.NewGuid():N}.txt");

        try
        {
            File.WriteAllText(unsupportedPath, "not office");
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));

            var missing = RunOfficeCli("get", missingPath, "/body", "--json");
            Assert.Equal(1, missing.ExitCode);
            AssertJsonFailure(missing, "file_not_found");

            var unsupported = RunOfficeCli("get", unsupportedPath, "/", "--json");
            Assert.Equal(1, unsupported.ExitCode);
            AssertJsonFailure(unsupported, "unsupported_type");

            var outOfRange = RunOfficeCli("get", path, "/body/p[99]", "--json");
            Assert.Equal(1, outOfRange.ExitCode);
            AssertJsonFailure(outOfRange, "not_found");

            var selected = RunOfficeCli("get", path, "selected", "--json");
            Assert.Equal(1, selected.ExitCode);
            using (var selectedDoc = JsonDocument.Parse(selected.Stdout))
            {
                Assert.False(selectedDoc.RootElement.GetProperty("success").GetBoolean());
                Assert.Contains("no watch running", selectedDoc.RootElement.GetProperty("message").GetString());
                Assert.False(selectedDoc.RootElement.TryGetProperty("error", out _));
            }

            var bareSet = RunOfficeCli("set", path, "paragraph", "--prop", "align=center", "--json");
            Assert.Equal(1, bareSet.ExitCode);
            AssertJsonFailure(bareSet, "bare_selector_rejected");

            var scopedSet = AssertJsonSuccess(RunOfficeCli("set", path, "/", "--prop", "title=scoped ok", "--json"));
            Assert.Contains("Updated /", scopedSet.GetString() ?? scopedSet.GetRawText());

            var bareRemove = RunOfficeCli("remove", path, "paragraph", "--json");
            Assert.Equal(1, bareRemove.ExitCode);
            AssertJsonFailure(bareRemove, "bare_selector_rejected");
        }
        finally
        {
            try { File.Delete(path); } catch { }
            try { File.Delete(unsupportedPath); } catch { }
        }
    }

    [Fact]
    public void WatchCommandFamily_TracksExplicitMarkAndParagraphGotoLifecycle()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_watch_contract_{Guid.NewGuid():N}.docx");
        var port = 26315 + Random.Shared.Next(1, 1000);
        var startInfo = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        startInfo.ArgumentList.Add("watch");
        startInfo.ArgumentList.Add(path);
        startInfo.ArgumentList.Add("--port");
        startInfo.ArgumentList.Add(port.ToString());
        startInfo.Environment["OFFICECLI_WATCH_IDLE_SECONDS"] = "60";
        startInfo.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";

        Process? watchProcess = null;
        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli(
                "add", path, "/body", "--type", "paragraph", "--prop", "text=watch target", "--json"));
            using (var document = new WordHandler(path, editable: true))
            {
                document.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
                document.Set("/body/tbl[1]/tr[1]", new() { ["c1"] = "watch table cell" });
                document.Add("/body", "paragraph", null, new() { ["text"] = "TODO-123" });
            }

            watchProcess = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start watch");
            _ = watchProcess.StandardOutput.ReadToEndAsync();
            _ = watchProcess.StandardError.ReadToEndAsync();
            Assert.True(
                SpinWait.SpinUntil(() => OfficeCli.Core.WatchServer.IsWatching(path), TimeSpan.FromSeconds(10)),
                $"watch did not start (exited={watchProcess.HasExited}).");

            var selected = AssertJsonSuccess(RunOfficeCli("get", path, "selected", "--json"));
            Assert.Equal(0, selected.GetProperty("matches").GetInt32());
            Assert.Empty(selected.GetProperty("results").EnumerateArray());

            var initialMarks = RunOfficeCli("get-marks", path, "--json");
            Assert.True(initialMarks.ExitCode == 0, initialMarks.Stderr);
            using (var initialMarksDocument = JsonDocument.Parse(initialMarks.Stdout))
            {
                Assert.Equal(0, initialMarksDocument.RootElement.GetProperty("version").GetInt32());
                Assert.Empty(initialMarksDocument.RootElement.GetProperty("marks").EnumerateArray());
            }

            var marked = RunOfficeCli(
                "mark", path, "/body/p[1]",
                "--prop", "note=review this", "--prop", "color=yellow", "--json");
            Assert.True(marked.ExitCode == 0, marked.Stderr);
            string markId;
            using (var markedDocument = JsonDocument.Parse(marked.Stdout))
            {
                var mark = markedDocument.RootElement;
                markId = mark.GetProperty("id").GetString()!;
                Assert.False(string.IsNullOrWhiteSpace(markId));
                Assert.Equal("/body/p[1]", mark.GetProperty("path").GetString());
                Assert.Equal("review this", mark.GetProperty("note").GetString());
            }

            var invalidColor = RunOfficeCli(
                "mark", path, "/body/p[1]",
                "--prop", "color=javascript:alert(1)", "--json");
            Assert.Equal(1, invalidColor.ExitCode);
            Assert.Contains("invalid color", invalidColor.Stdout, StringComparison.OrdinalIgnoreCase);

            var regexMarked = RunOfficeCli(
                "mark", path, "/body/p[2]",
                "--prop", "find=TODO-[0-9]+", "--prop", "regex=true",
                "--prop", "color=00ff00", "--prop", "note=regex mark", "--json");
            Assert.True(regexMarked.ExitCode == 0, regexMarked.Stderr);
            using (var regexDocument = JsonDocument.Parse(regexMarked.Stdout))
            {
                var regexMark = regexDocument.RootElement;
                Assert.Equal("/body/p[2]", regexMark.GetProperty("path").GetString());
                Assert.Equal("r\"TODO-[0-9]+\"", regexMark.GetProperty("find").GetString());
                Assert.Equal("#00FF00", regexMark.GetProperty("color").GetString());
                Assert.Contains(regexMark.GetProperty("matched_text").EnumerateArray(),
                    item => item.GetString() == "TODO-123");
            }

            var listed = RunOfficeCli("get-marks", path, "--json");
            Assert.True(listed.ExitCode == 0, listed.Stderr);
            using (var listedDocument = JsonDocument.Parse(listed.Stdout))
            {
                var listedMarks = listedDocument.RootElement.GetProperty("marks").EnumerateArray().ToArray();
                Assert.Equal(2, listedMarks.Length);
                Assert.Contains(listedMarks, mark => mark.GetProperty("id").GetString() == markId);
                Assert.Contains(listedMarks, mark => mark.GetProperty("path").GetString() == "/body/p[2]");
                Assert.True(listedDocument.RootElement.GetProperty("version").GetInt32() >= 2);
            }

            var gotoResult = AssertJsonSuccess(RunOfficeCli(
                "goto", path, "/body/p[1]", "--json"));
            Assert.Contains("/body/p[1]", gotoResult.GetString() ?? gotoResult.GetRawText());

            foreach (var tablePath in new[]
            {
                "/body/table[1]",
                "/body/table[1]/tr[1]",
                "/body/table[1]/tr[1]/tc[1]",
            })
            {
                var tableGoto = AssertJsonSuccess(RunOfficeCli("goto", path, tablePath, "--json"));
                Assert.Contains(tablePath, tableGoto.GetString() ?? tableGoto.GetRawText());
            }

            var missingGoto = RunOfficeCli("goto", path, "/body/p[99]", "--json");
            Assert.Equal(1, missingGoto.ExitCode);
            using (var missingGotoDocument = JsonDocument.Parse(missingGoto.Stdout))
            {
                Assert.False(missingGotoDocument.RootElement.GetProperty("success").GetBoolean());
                Assert.Contains("Cannot scroll", missingGotoDocument.RootElement.GetProperty("message").GetString());
            }

            var conflictingUnmark = RunOfficeCli(
                "unmark", path, "--path", "/body/p[1]", "--all", "--json");
            Assert.Equal(2, conflictingUnmark.ExitCode);
            Assert.Contains("either --path or --all", conflictingUnmark.Stdout);

            var unmarked = AssertJsonSuccess(RunOfficeCli(
                "unmark", path, "--path", "/body/p[1]", "--json"));
            Assert.Contains("Removed 1 mark", unmarked.GetString() ?? unmarked.GetRawText());

            var allUnmarked = AssertJsonSuccess(RunOfficeCli(
                "unmark", path, "--all", "--json"));
            Assert.Contains("Removed 1 mark", allUnmarked.GetString() ?? allUnmarked.GetRawText());

            var finalMarks = RunOfficeCli("get-marks", path, "--json");
            Assert.True(finalMarks.ExitCode == 0, finalMarks.Stderr);
            using var finalMarksDocument = JsonDocument.Parse(finalMarks.Stdout);
            Assert.Empty(finalMarksDocument.RootElement.GetProperty("marks").EnumerateArray());
        }
        finally
        {
            try { RunOfficeCli("unwatch", path); } catch { }
            if (watchProcess != null)
            {
                if (!watchProcess.WaitForExit(5_000))
                {
                    try { watchProcess.Kill(entireProcessTree: true); } catch { }
                }
                watchProcess.Dispose();
            }
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void SwapCli_SwapsSameParentElementsAndRejectsInvalidPairs()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_swap_{Guid.NewGuid():N}.docx");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", path, "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--prop", "text=A", "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--prop", "text=B", "--json"));
            AssertJsonSuccess(RunOfficeCli("add", path, "/body", "--type", "paragraph", "--prop", "text=C", "--json"));
            using (var handler = new WordHandler(path, editable: true))
            {
                handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });
                handler.Set("/body/tbl[1]/tr[1]", new() { ["c1"] = "left", ["c2"] = "right" });
            }

            var bodySwap = AssertJsonSuccess(RunOfficeCli("swap", path, "/body/p[1]", "/body/p[3]", "--json"));
            Assert.Contains("Swapped", bodySwap.GetString() ?? bodySwap.GetRawText());

            var cellSwap = AssertJsonSuccess(RunOfficeCli(
                "swap", path, "/body/tbl[1]/tr[1]/tc[1]", "/body/tbl[1]/tr[1]/tc[2]", "--json"));
            Assert.Contains("Swapped", cellSwap.GetString() ?? cellSwap.GetRawText());

            using (var handler = new WordHandler(path, editable: false))
            {
                Assert.Equal(["C", "B", "A"], handler.Query("paragraph").Select(p => p.Text ?? "").Take(3).ToArray());
                Assert.Equal("right", handler.Get("/body/tbl[1]/tr[1]/tc[1]").Text);
                Assert.Equal("left", handler.Get("/body/tbl[1]/tr[1]/tc[2]").Text);
            }

            var crossParent = RunOfficeCli("swap", path, "/body/p[1]", "/body/tbl[1]/tr[1]/tc[1]", "--json");
            Assert.Equal(1, crossParent.ExitCode);
            AssertJsonFailure(crossParent, "invalid_value");
            Assert.Contains("different parents", crossParent.Stdout);

            var missing = RunOfficeCli("swap", path, "/body/p[99]", "/body/p[1]", "--json");
            Assert.Equal(1, missing.ExitCode);
            AssertJsonFailure(missing, "not_found");
            Assert.Contains("Element not found", missing.Stdout);

            AssertJsonSuccess(RunOfficeCli("validate", path, "--json"));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void MergeCli_ReplacesDocxBodyTableHeaderFooterAndReportsUnresolved()
    {
        var templatePath = Path.Combine(Path.GetTempPath(), $"officecli_merge_template_{Guid.NewGuid():N}.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), $"officecli_merge_output_{Guid.NewGuid():N}.docx");
        var dataPath = Path.Combine(Path.GetTempPath(), $"officecli_merge_data_{Guid.NewGuid():N}.json");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", templatePath, "--json"));
            using (var handler = new WordHandler(templatePath, editable: true))
            {
                handler.Add("/body", "paragraph", null, new()
                {
                    ["text"] = "Hello {{name}} / {{ user.title }} / {{literal.key}} / {{missing}}"
                });
                handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
                handler.Set("/body/tbl[1]/tr[1]", new() { ["c1"] = "Table {{items[0]}}" });
                handler.Add("/", "header", null, new() { ["text"] = "Header {{name}}" });
                handler.Add("/", "footer", null, new() { ["text"] = "Footer {{items[1]}}" });
            }

            File.WriteAllText(dataPath, """
                {
                  "name": "Ada",
                  "user": { "title": "Dr" },
                  "items": ["alpha", "beta"],
                  "literal.key": "literal wins"
                }
                """);

            var mergeData = AssertJsonSuccess(RunOfficeCli(
                "merge", templatePath, outputPath, "--data", dataPath, "--json"));
            Assert.Equal(Path.GetFullPath(outputPath), mergeData.GetProperty("output").GetString());
            Assert.True(mergeData.GetProperty("replacedKeys").GetInt32() >= 5);
            Assert.Contains("missing", mergeData.GetProperty("unresolvedPlaceholders").EnumerateArray().Select(x => x.GetString()));

            using (var handler = new WordHandler(outputPath, editable: false))
            {
                Assert.Contains("Hello Ada / Dr / literal wins / {{missing}}", handler.Raw("/document"));
                Assert.Contains("Table alpha", handler.Raw("/document"));
                Assert.Contains("Header Ada", handler.Raw("/header[1]"));
                Assert.Contains("Footer beta", handler.Raw("/footer[1]"));
            }

            var duplicate = RunOfficeCli("merge", templatePath, outputPath, "--data", dataPath, "--json");
            Assert.Equal(1, duplicate.ExitCode);
            AssertJsonFailure(duplicate, "file_exists");

            var invalidJson = RunOfficeCli(
                "merge", templatePath, outputPath, "--data", "[]", "--force", "--json");
            Assert.Equal(1, invalidJson.ExitCode);
            AssertJsonFailure(invalidJson, "invalid_json");
        }
        finally
        {
            try { File.Delete(templatePath); } catch { }
            try { File.Delete(outputPath); } catch { }
            try { File.Delete(dataPath); } catch { }
        }
    }

    [Fact]
    public void MergeCli_ReplacesNotesAndCommentsButKeepsCrossRunPlaceholderUnresolved()
    {
        var templatePath = Path.Combine(Path.GetTempPath(), $"officecli_merge_parts_{Guid.NewGuid():N}.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), $"officecli_merge_parts_output_{Guid.NewGuid():N}.docx");
        var dataPath = Path.Combine(Path.GetTempPath(), $"officecli_merge_parts_data_{Guid.NewGuid():N}.json");

        try
        {
            AssertJsonSuccess(RunOfficeCli("create", templatePath, "--json"));
            using (var handler = new WordHandler(templatePath, editable: true))
            {
                var splitParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
                handler.Add(splitParagraph, "run", null, new() { ["text"] = "{{na" });
                handler.Add(splitParagraph, "run", null, new() { ["text"] = "me}}" });

                var footnoteHost = handler.Add("/body", "paragraph", null, new() { ["text"] = "Footnote host" });
                handler.Add(footnoteHost, "footnote", null, new() { ["text"] = "{{foot}}" });

                var endnoteHost = handler.Add("/body", "paragraph", null, new() { ["text"] = "Endnote host" });
                handler.Add(endnoteHost, "endnote", null, new() { ["text"] = "{{end}}" });

                var commentHost = handler.Add("/body", "paragraph", null, new() { ["text"] = "Comment host" });
                handler.Add(commentHost, "comment", null, new() { ["text"] = "{{comment}}" });
            }

            File.WriteAllText(dataPath, """
                {
                  "foot": "foot Ada",
                  "end": "end Ada",
                  "comment": "comment Ada"
                }
                """);

            var mergeData = AssertJsonSuccess(RunOfficeCli(
                "merge", templatePath, outputPath, "--data", dataPath, "--json"));
            Assert.Equal(3, mergeData.GetProperty("replacedKeys").GetInt32());
            Assert.Contains("name", mergeData.GetProperty("unresolvedPlaceholders")
                .EnumerateArray().Select(value => value.GetString()));

            using var merged = new WordHandler(outputPath, editable: false);
            Assert.Contains("{{na", merged.Raw("/document"));
            Assert.Equal("foot Ada", Assert.Single(merged.Query("footnote")).Text);
            Assert.Equal("end Ada", Assert.Single(merged.Query("endnote")).Text);
            Assert.Equal("comment Ada", Assert.Single(merged.Query("comment")).Text);
        }
        finally
        {
            try { File.Delete(templatePath); } catch { }
            try { File.Delete(outputPath); } catch { }
            try { File.Delete(dataPath); } catch { }
        }
    }

    [Fact]
    public void CommentQueryCli_DirectionSelectorReportsCurrentUnsupportedSurface()
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
        {
            var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Comment direction host" });
            handler.Add(paragraphPath, "comment", null, new()
            {
                ["text"] = "RTL review",
                ["author"] = "Reviewer",
                ["direction"] = "rtl"
            });
        }

        var all = AssertSingleResult(AssertJsonSuccess(RunOfficeCli(
            "query", path, "comment", "--json")));
        Assert.Equal("RTL review", all.GetProperty("text").GetString());
        Assert.False(all.GetProperty("format").TryGetProperty("direction", out _));

        var filteredResult = RunOfficeCli(
            "query", path, "comment[direction=rtl]", "--json");
        var filtered = AssertJsonSuccess(filteredResult);
        Assert.Equal(0, filtered.GetProperty("matches").GetInt32());
        using var filteredDocument = JsonDocument.Parse(filteredResult.Stdout);
        Assert.Contains(filteredDocument.RootElement.GetProperty("warnings").EnumerateArray(), warning =>
            warning.GetProperty("key").GetString() == "direction"
            && warning.GetProperty("kind").GetString() == "unknown_key");
    }

    [Fact]
    public void HelpCli_DocxRevisionSchemaMatchesActionAndReadbackSurface()
    {
        static JsonElement ReadSchema(CliRunResult result)
        {
            Assert.Equal(0, result.ExitCode);
            using var document = JsonDocument.Parse(result.Stdout);
            return document.RootElement.Clone();
        }

        var revision = ReadSchema(RunOfficeCli("help", "docx", "revision", "--json"));
        Assert.Equal("docx", revision.GetProperty("format").GetString());
        Assert.Equal("revision", revision.GetProperty("element").GetString());
        Assert.False(revision.GetProperty("operations").GetProperty("add").GetBoolean());
        Assert.True(revision.GetProperty("operations").GetProperty("set").GetBoolean());
        Assert.True(revision.GetProperty("operations").GetProperty("query").GetBoolean());

        var properties = revision.GetProperty("properties");
        Assert.True(properties.GetProperty("action").GetProperty("set").GetBoolean());
        Assert.True(properties.GetProperty("nativePath").GetProperty("get").GetBoolean());
        Assert.False(properties.GetProperty("author").GetProperty("add").GetBoolean());

        var setRevision = ReadSchema(RunOfficeCli("help", "docx", "set", "revision", "--json"));
        Assert.Equal("revision", setRevision.GetProperty("element").GetString());

        var all = RunOfficeCli("help", "docx", "all", "--json");
        Assert.Equal(0, all.ExitCode);
        using var allDocument = JsonDocument.Parse(all.Stdout);
        var allData = allDocument.RootElement.GetProperty("data");
        Assert.Contains(allData.EnumerateArray(), item =>
            item.GetProperty("kind").GetString() == "ELEM"
            && item.GetProperty("element").GetString() == "revision"
            && item.GetProperty("ops").GetString()!.Contains('s'));
    }

    [Fact]
    public void HelpCli_DocxAllListsEveryEmbeddedSchemaElementWithOperations()
    {
        var result = RunOfficeCli("help", "docx", "all", "--json");
        Assert.Equal(0, result.ExitCode);
        using var document = JsonDocument.Parse(result.Stdout);
        var data = document.RootElement.GetProperty("data");
        var elements = data.EnumerateArray()
            .Where(item => item.GetProperty("kind").GetString() == "ELEM")
            .ToList();

        var repo = new DirectoryInfo(Path.GetDirectoryName(FindOfficeCliDll())!);
        while (repo != null && !Directory.Exists(Path.Combine(repo.FullName, "schemas", "help", "docx")))
            repo = repo.Parent!;
        Assert.NotNull(repo);

        var schemaNames = Directory.GetFiles(
                Path.Combine(repo!.FullName, "schemas", "help", "docx"), "*.json")
            .Select(file => Path.GetFileNameWithoutExtension(file))
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
        var helpNames = elements
            .Select(item => item.GetProperty("element").GetString()!)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(schemaNames, helpNames);
        Assert.All(elements, item =>
        {
            var ops = item.GetProperty("ops").GetString();
            Assert.NotNull(ops);
            Assert.Equal(5, ops!.Length);
            if (item.GetProperty("element").GetString() != "raw")
                Assert.True(item.GetProperty("paths").GetArrayLength() > 0);
        });
    }

    [Fact]
    public void ViewJson_CorruptDocxReturnsErrorEnvelopeAndLeavesBytesUntouched()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_corrupt_{Guid.NewGuid():N}.docx");
        var original = new byte[] { 0x6E, 0x6F, 0x74, 0x2D, 0x64, 0x6F, 0x63, 0x78 };
        File.WriteAllBytes(path, original);

        try
        {
            var result = RunOfficeCli("view", path, "text", "--json");

            Assert.Equal(1, result.ExitCode);
            using var doc = JsonDocument.Parse(result.Stdout);
            var root = doc.RootElement;
            Assert.False(root.GetProperty("success").GetBoolean());
            Assert.Equal("corrupt_file", root.GetProperty("error").GetProperty("code").GetString());
            Assert.Equal(original, File.ReadAllBytes(path));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void RefreshCli_ReportsUnsupportedTypeAndPlatformFallbackOutcome()
    {
        var path = CreateBlankDocx();
        var unsupportedPath = Path.Combine(Path.GetTempPath(), $"officecli_refresh_{Guid.NewGuid():N}.txt");
        File.WriteAllText(unsupportedPath, "not a document");

        try
        {
            var unsupported = RunOfficeCli("refresh", unsupportedPath, "--json");
            Assert.Equal(1, unsupported.ExitCode);
            AssertJsonFailure(unsupported, "unsupported_type");

            var refresh = RunOfficeCli("refresh", path, "--json");
            if (refresh.ExitCode == 0)
            {
                var data = AssertJsonSuccess(refresh);
                Assert.Contains("Refreshed:", data.GetString() ?? data.GetRawText());
                Assert.Contains("backend:", data.GetString() ?? data.GetRawText());
            }
            else
            {
                AssertJsonFailure(refresh, "refresh_failed");
                Assert.Contains("backend unavailable", refresh.Stdout, StringComparison.OrdinalIgnoreCase);
            }
        }
        finally
        {
            try { File.Delete(unsupportedPath); } catch { }
        }
    }

    [Fact]
    public void RefreshCli_FieldFixtureRecordsBackendAndPostRefreshPackageState()
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Refresh heading" });
            handler.Add("/body", "toc", null, new() { ["levels"] = "1-2" });
            var pageParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "Page: " });
            handler.Add(pageParagraph, "field", null, new() { ["fieldType"] = "page", ["text"] = "" });
        }

        var refresh = RunOfficeCli("refresh", path, "--json");
        if (refresh.ExitCode != 0)
        {
            AssertJsonFailure(refresh, "refresh_failed");
            Assert.Contains("backend unavailable", refresh.Stdout, StringComparison.OrdinalIgnoreCase);
            return;
        }

        var data = AssertJsonSuccess(refresh);
        var message = data.GetString() ?? data.GetRawText();
        Assert.Contains("Refreshed:", message);
        Assert.Contains("backend:", message);

        using var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, false);
        var properties = document.ExtendedFilePropertiesPart?.Properties;
        Assert.NotNull(properties);
        Assert.False(string.IsNullOrWhiteSpace(properties!.Pages?.Text));
    }

    [Fact]
    public async Task McpCli_WordToolRunsCreateAddQueryWithJsonRpcEnvelope()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_mcp_word_{Guid.NewGuid():N}.docx");
        var startInfo = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("mcp");
        startInfo.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start MCP server");

        async Task<JsonDocument> Send(string request)
        {
            await process.StandardInput.WriteLineAsync(request);
            await process.StandardInput.FlushAsync();
            var line = await process.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(5));
            Assert.NotNull(line);
            return JsonDocument.Parse(line!);
        }

        string ToolCall(int id, params string[] argv)
        {
            return JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                id,
                method = "tools/call",
                @params = new
                {
                    name = "officecli",
                    arguments = new { command = argv }
                }
            });
        }

        try
        {
            using (var initialize = await Send("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{}}"))
            {
                Assert.Equal("2024-11-05", initialize.RootElement
                    .GetProperty("result").GetProperty("protocolVersion").GetString());
            }

            using (var tools = await Send("{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"tools/list\",\"params\":{}}"))
            {
                Assert.Equal("officecli", tools.RootElement.GetProperty("result")
                    .GetProperty("tools")[0].GetProperty("name").GetString());
            }

            using (var create = await Send(ToolCall(3, "create", path, "--json")))
            {
                Assert.False(create.RootElement.GetProperty("result").GetProperty("isError").GetBoolean());
                var text = create.RootElement.GetProperty("result").GetProperty("content")[0]
                    .GetProperty("text").GetString();
                using var envelope = JsonDocument.Parse(text!);
                Assert.True(envelope.RootElement.GetProperty("success").GetBoolean());
            }

            using (var add = await Send(ToolCall(4, "add", path, "/body", "--type", "paragraph", "--prop", "text=MCP Word", "--json")))
            {
                Assert.False(add.RootElement.GetProperty("result").GetProperty("isError").GetBoolean());
                var text = add.RootElement.GetProperty("result").GetProperty("content")[0]
                    .GetProperty("text").GetString();
                using var envelope = JsonDocument.Parse(text!);
                Assert.True(envelope.RootElement.GetProperty("success").GetBoolean());
            }

            using (var query = await Send(ToolCall(5, "query", path, "paragraph", "--json")))
            {
                Assert.False(query.RootElement.GetProperty("result").GetProperty("isError").GetBoolean());
                var text = query.RootElement.GetProperty("result").GetProperty("content")[0]
                    .GetProperty("text").GetString();
                using var envelope = JsonDocument.Parse(text!);
                Assert.Equal(1, envelope.RootElement.GetProperty("data").GetProperty("matches").GetInt32());
                Assert.Equal("MCP Word", envelope.RootElement.GetProperty("data").GetProperty("results")[0]
                    .GetProperty("text").GetString());
            }

            using (var unknown = await Send(JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                id = 6,
                method = "tools/call",
                @params = new { name = "unknown", arguments = new { command = new[] { "help" } } }
            })))
            {
                var error = unknown.RootElement.GetProperty("error");
                Assert.Equal(-32602, error.GetProperty("code").GetInt32());
                Assert.Contains("Unknown tool", error.GetProperty("message").GetString());
            }
        }
        finally
        {
            try { process.StandardInput.Close(); } catch { }
            if (!process.WaitForExit(5_000))
            {
                try { process.Kill(entireProcessTree: true); } catch { }
            }
            try { File.Delete(path); } catch { }
        }
    }

    private static JsonElement AssertJsonSuccess(CliRunResult result, int expectedExitCode = 0)
    {
        Assert.True(result.ExitCode == expectedExitCode, $"exit={result.ExitCode}\nstdout:\n{result.Stdout}\nstderr:\n{result.Stderr}");
        using var doc = JsonDocument.Parse(result.Stdout);
        var root = doc.RootElement.Clone();
        Assert.True(root.GetProperty("success").GetBoolean(), result.Stdout);
        return root.GetProperty("data").Clone();
    }

    private static void AssertJsonFailure(CliRunResult result, string code)
    {
        using var doc = JsonDocument.Parse(result.Stdout);
        var root = doc.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean(), result.Stdout);
        Assert.Equal(code, root.GetProperty("error").GetProperty("code").GetString());
    }

    private static void AssertJsonWarningFailure(CliRunResult result, string code)
    {
        using var doc = JsonDocument.Parse(result.Stdout);
        var root = doc.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean(), result.Stdout);
        AssertJsonWarning(root, code);
    }

    private static void AssertJsonWarning(CliRunResult result, string code)
    {
        using var doc = JsonDocument.Parse(result.Stdout);
        AssertJsonWarning(doc.RootElement, code);
    }

    private static void AssertJsonWarning(JsonElement root, string code)
    {
        Assert.Contains(root.GetProperty("warnings").EnumerateArray(),
            warning => warning.GetProperty("code").GetString() == code);
    }

    private static void AssertProtectedFailure(CliRunResult result)
    {
        Assert.Equal(1, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        Assert.False(doc.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
        Assert.Contains("Document is protected", result.Stdout);
    }

    private static JsonElement AssertSingleResult(JsonElement data)
    {
        Assert.Equal(1, data.GetProperty("matches").GetInt32());
        return data.GetProperty("results")[0].Clone();
    }

    private static void AssertBatchSourceSuccess(CliRunResult result, int expectedTotal = 2)
    {
        var data = AssertJsonSuccess(result);
        var summary = data.GetProperty("summary");
        Assert.Equal(expectedTotal, summary.GetProperty("total").GetInt32());
        Assert.Equal(expectedTotal, summary.GetProperty("succeeded").GetInt32());
        Assert.Equal(0, summary.GetProperty("failed").GetInt32());
    }

    private static CliRunResult RunOfficeCli(params string[] args) => RunOfficeCliCore(null, args);

    private static CliRunResult RunOfficeCliWithInput(string input, params string[] args) =>
        RunOfficeCliCore(input, args);

    private static CliRunResult RunOfficeCliCore(string? input, params string[] args)
    {
        var officeCliDll = FindOfficeCliDll();
        var psi = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            RedirectStandardInput = input != null,
        };
        psi.ArgumentList.Add(officeCliDll);
        foreach (var arg in args) psi.ArgumentList.Add(arg);
        psi.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";

        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start officecli");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (input != null)
        {
            process.StandardInput.Write(input);
            process.StandardInput.Close();
        }

        if (!process.WaitForExit(30_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("officecli process did not exit within 30 seconds");
        }

        return new CliRunResult(process.ExitCode, stdoutTask.GetAwaiter().GetResult(), stderrTask.GetAwaiter().GetResult());
    }

    private static CliRunResult RunOfficeCliAppHost(params string[] args)
    {
        var psi = new ProcessStartInfo(FindOfficeCliAppHost())
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        foreach (var arg in args) psi.ArgumentList.Add(arg);
        psi.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";

        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start officecli apphost");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        if (!process.WaitForExit(30_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("officecli apphost process did not exit within 30 seconds");
        }

        return new CliRunResult(process.ExitCode, stdoutTask.GetAwaiter().GetResult(), stderrTask.GetAwaiter().GetResult());
    }

    private static string FindOfficeCliDll()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "officecli.slnx")))
        {
            dir = dir.Parent;
        }

        if (dir == null)
            throw new InvalidOperationException("Could not locate officecli.slnx from test output directory");

        var candidates = Directory.GetFiles(
                Path.Combine(dir.FullName, "src", "officecli", "bin"),
                "officecli.dll",
                SearchOption.AllDirectories)
            .Where(path => File.Exists(Path.Combine(Path.GetDirectoryName(path)!, "officecli.runtimeconfig.json")))
            .OrderByDescending(IsCurrentHostRuntimePath)
            .ThenByDescending(File.GetLastWriteTimeUtc)
            .ToList();

        return candidates.FirstOrDefault()
            ?? throw new InvalidOperationException("Could not locate built officecli.dll under src/officecli/bin");
    }

    private static bool IsCurrentHostRuntimePath(string path)
    {
        var normalized = path.Replace('\\', '/');
        var marker = OperatingSystem.IsWindows()
            ? "/win-"
            : OperatingSystem.IsMacOS()
                ? "/osx-"
                : "/linux-";
        return normalized.Contains(marker, StringComparison.OrdinalIgnoreCase);
    }

    private static string FindOfficeCliAppHost()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "officecli.slnx")))
        {
            dir = dir.Parent;
        }

        if (dir == null)
            throw new InvalidOperationException("Could not locate officecli.slnx from test output directory");

        var executableName = OperatingSystem.IsWindows() ? "officecli.exe" : "officecli";
        var candidates = Directory.GetFiles(
                Path.Combine(dir.FullName, "src", "officecli", "bin"),
                executableName,
                SearchOption.AllDirectories)
            .Where(path => new FileInfo(path).Length > 0)
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .ToList();

        return candidates.FirstOrDefault()
            ?? throw new InvalidOperationException("Could not locate built officecli apphost under src/officecli/bin");
    }

    private sealed record CliRunResult(int ExitCode, string Stdout, string Stderr);
}
