using System.Text.Json;
using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Collection("Word CLI output")]
[Trait("Speed", "Integration")]
public sealed class WordAddTypeCommandIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void AddCli_BlockTableAndNoteRoutes_ReturnReadablePathsAndValidate()
    {
        var path = CreateBlankDocx();
        var paragraph = Add(path, "/body", "paragraph", "text=block host");
        var table = Add(path, "/body", "table", "rows=1", "cols=1");
        var row = Add(path, table, "row", "cols=2", "c1=row one", "c2=row two");
        var column = Add(path, table, "col", "width=1600", "text=column");
        Add(path, row, "cell", "text=added cell", "width=1400");
        var section = Add(path, "/body", "section", "type=nextPage");
        Add(path, "/", "header", "type=default", "text=CLI header");
        Add(path, "/", "footer", "type=default", "text=CLI footer");
        Add(path, paragraph, "footnote", "text=CLI footnote");
        Add(path, paragraph, "endnote", "text=CLI endnote");
        Add(path, paragraph, "comment", "text=CLI comment", "author=Tester", "initials=TS");
        Add(path, "/body", "toc", "levels=1-2", "title=Contents");

        Assert.Equal(1, QueryCount(path, "table"));
        Assert.True(QueryCount(path, "row") >= 2);
        Assert.True(QueryCount(path, "cell") >= 3);
        Assert.Contains("/col[", column, StringComparison.Ordinal);
        Assert.Contains("/section[", section, StringComparison.Ordinal);
        Assert.Equal(1, QueryCount(path, "header"));
        Assert.Equal(1, QueryCount(path, "footer"));
        Assert.Equal(1, QueryCount(path, "footnote"));
        Assert.Equal(1, QueryCount(path, "endnote"));
        Assert.Equal(1, QueryCount(path, "comment"));
        Assert.Equal(1, QueryCount(path, "toc"));
        AssertJsonSuccess(InvokeCaptured("validate", path, "--json"));
    }

    [Fact]
    public void AddCli_InlineReferenceAndControlRoutes_ReturnReadablePathsAndValidate()
    {
        var path = CreateBlankDocx();
        var paragraph = Add(path, "/body", "p", "text=inline host");
        Add(path, paragraph, "r", "text=run");
        Add(path, paragraph, "tab", "pos=1440", "val=right", "leader=dot");
        Add(path, paragraph, "ptab", "align=right", "relativeTo=margin", "leader=dot");
        Add(path, paragraph, "field", "fieldType=page", "text=1");
        Add(path, paragraph, "link", "url=https://example.com", "text=link");
        Add(path, paragraph, "bookmark", "name=CliRouteBookmark");
        Add(path, paragraph, "pagebreak", "breakType=page");
        Add(path, "/body", "contentcontrol", "type=text", "text=CLI SDT");
        Add(path, "/body", "formfield", "name=CliForm", "type=text", "text=CLI form");
        Add(path, paragraph, "permstart", "id=17", "edGrp=everyone");
        Add(path, paragraph, "permend", "id=17");

        Assert.True(QueryCount(path, "paragraph") >= 1);
        Assert.True(QueryCount(path, "run") >= 1);
        Assert.True(QueryCount(path, "field") >= 1);
        Assert.Equal(1, QueryCount(path, "hyperlink"));
        Assert.True(QueryCount(path, "bookmark") >= 1);
        Assert.Equal(1, QueryCount(path, "sdt"));
        Assert.Equal(1, QueryCount(path, "formfield"));
        Assert.Equal(1, QueryCount(path, "permStart"));
        Assert.Equal(1, QueryCount(path, "permEnd"));
        AssertJsonSuccess(InvokeCaptured("validate", path, "--json"));
    }

    [Fact]
    public void AddCli_MediaAndDrawingRoutes_ReturnReadablePathsAndValidate()
    {
        var path = CreateBlankDocx();
        Add(path, "/body", "picture", $"src={TinyPngDataUri}", "alt=CLI picture");
        Add(path, "/body", "ole",
            "src=data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,SGVsbG8=",
            "oleKind=package",
            "contentType=application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "embedExt=docx", "progId=Word.Document.12", "display=icon", "name=Embedded Word");
        Add(path, "/body", "chart", "chartType=column", "title=CLI chart",
            "categories=Q1,Q2", "data=Revenue:10,20");
        Add(path, "/body", "equation", "formula=x^2", "mode=display");
        var diagram = Add(path, "/body", "diagram", "render=native", "mermaid=flowchart TD; A[Start] --> B[Done]");
        var shape = Add(path, "/body", "shape", "geometry=roundRect", "width=4cm", "height=2cm", "fill=EAF2FF");
        var textbox = Add(path, "/body", "textbox", "text=CLI textbox", "width=4cm", "height=2cm");
        Add(path, "/body", "watermark", "text=DRAFT", "rotation=-45", "opacity=.5");

        Assert.True(QueryCount(path, "picture") >= 1);
        Assert.Equal(1, QueryCount(path, "ole"));
        Assert.Equal(1, QueryCount(path, "chart"));
        Assert.Equal(1, QueryCount(path, "equation"));
        Assert.Contains("/group[", diagram, StringComparison.Ordinal);
        Assert.Contains("/shape[", shape, StringComparison.Ordinal);
        Assert.Contains("/textbox[", textbox, StringComparison.Ordinal);
        Assert.Equal(1, QueryCount(path, "watermark"));
        AssertJsonSuccess(InvokeCaptured("validate", path, "--json"));
    }

    [Fact]
    public void AddCli_DefinitionRoutes_CreateStylesAndNumberingGraph()
    {
        var path = CreateBlankDocx();
        Add(path, "/styles", "style", "id=CliRouteStyle", "name=CLI Route Style", "type=paragraph");
        var abstractNum = Add(path, "/numbering", "abstractNum", "id=77", "type=multilevel",
            "level0.format=decimal", "level0.text=%1.", "level0.indent=720");
        Add(path, "/numbering", "num", "id=33", "abstractNumId=77", "start=4");
        var level = Add(path, abstractNum, "lvl", "ilvl=1", "format=lowerLetter", "text=%2.", "indent=1440");

        Assert.True(QueryCount(path, "style") >= 1);
        Assert.Equal(1, QueryCount(path, "abstractNum"));
        Assert.Equal(1, QueryCount(path, "num"));
        Assert.Contains("/lvl[", level, StringComparison.Ordinal);
        AssertJsonSuccess(InvokeCaptured("validate", path, "--json"));
    }

    [Fact]
    public void AddCli_AliasesAndUnsafeDirectTypes_FollowCurrentProtocol()
    {
        var path = CreateBlankDocx();
        var paragraph = Add(path, "/body", "p", "text=alias paragraph");
        Add(path, paragraph, "r", "text=alias run");
        Add(path, "/body", "tbl", "data=A,B");
        Add(path, "/body", "img", $"src={TinyPngDataUri}");

        var before = ReadDocumentXml(path);
        foreach (var rejectedType in new[] { "ins", "del", "moveTo", "moveFrom", "altChunk" })
        {
            var result = InvokeCaptured("add", path, "/body", "--type", rejectedType, "--json");
            Assert.NotEqual(0, result.ExitCode);
            using var json = JsonDocument.Parse(result.Stdout);
            Assert.False(json.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
        }

        Assert.Equal(before, ReadDocumentXml(path));
        AssertJsonSuccess(InvokeCaptured("validate", path, "--json"));
    }

    private static string Add(string path, string parent, string type, params string[] props)
    {
        var args = new List<string> { "add", path, parent, "--type", type };
        foreach (var prop in props)
        {
            args.Add("--prop");
            args.Add(prop);
        }
        args.Add("--json");

        var result = InvokeCaptured(args.ToArray());
        Assert.True(result.ExitCode is 0 or 2,
            $"add {type} failed with exit {result.ExitCode}.\nstdout:\n{result.Stdout}\nstderr:\n{result.Stderr}");
        using var json = JsonDocument.Parse(result.Stdout);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
        var message = json.RootElement.GetProperty("data").GetString() ?? string.Empty;
        var marker = message.IndexOf(" at ", StringComparison.Ordinal);
        Assert.True(marker >= 0, message);
        var pathEnd = message.IndexOf('\n', marker + 4);
        return message[(marker + 4)..(pathEnd >= 0 ? pathEnd : message.Length)].Trim();
    }

    private static int QueryCount(string path, string selector)
    {
        var result = InvokeCaptured("query", path, selector, "--json");
        AssertJsonSuccess(result);
        using var json = JsonDocument.Parse(result.Stdout);
        return json.RootElement.GetProperty("data").GetProperty("matches").GetInt32();
    }

    private static string ReadDocumentXml(string path)
    {
        using var handler = new WordHandler(path, editable: false);
        return handler.Raw("/document");
    }

    private static void AssertJsonSuccess(CommandResult result)
    {
        Assert.Equal(0, result.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(result.Stderr), result.Stderr);
        using var json = JsonDocument.Parse(result.Stdout);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
    }

    private static CommandResult InvokeCaptured(params string[] args)
    {
        lock (typeof(WordAddTypeCommandIntegrationTests))
        {
            var oldOut = Console.Out;
            var oldError = Console.Error;
            using var stdout = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
            using var stderr = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
            Console.SetOut(stdout);
            Console.SetError(stderr);
            try
            {
                var root = CommandBuilder.BuildRootCommand();
                var exitCode = root.Parse(args).Invoke();
                return new CommandResult(exitCode, stdout.ToString(), stderr.ToString());
            }
            finally
            {
                Console.SetOut(oldOut);
                Console.SetError(oldError);
            }
        }
    }

    private sealed record CommandResult(int ExitCode, string Stdout, string Stderr);
}
