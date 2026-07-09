using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public class WordBatchTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void ExecuteBatch_AppliesMultiStepWordMutation()
    {
        var path = CreateBlankDocx();
        const string itemsJson = """
        [
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Batch paragraph"}},
          {"command":"set","path":"/body/p[1]","props":{"align":"center"}},
          {"command":"get","path":"/body/p[1]"}
        ]
        """;

        using var handler = new WordHandler(path, editable: true);
        var output = BatchExecutor.ExecuteBatch(handler, itemsJson, json: false);

        Assert.Contains("Batch complete: 3 succeeded, 0 failed, 3 total", output);
        var paragraph = handler.Get("/body/p[1]");
        Assert.Equal("Batch paragraph", paragraph.Text);
        Assert.Equal("center", Fmt(paragraph)["align"]);
    }

    [Fact]
    public void ExecuteBatch_DefaultContinuesAfterFailedItem()
    {
        var path = CreateBlankDocx();
        const string itemsJson = """
        [
          {"command":"add","parent":"/body/p[99]","type":"paragraph","props":{"text":"bad"}},
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"good"}}
        ]
        """;

        using var handler = new WordHandler(path, editable: true);
        var output = BatchExecutor.ExecuteBatch(handler, itemsJson, json: false);

        Assert.Contains("[1] ERROR:", output);
        Assert.Contains("Batch complete: 1 succeeded, 1 failed, 2 total", output);
        Assert.Empty(handler.Query("paragraph:contains(\"bad\")"));
        Assert.Single(handler.Query("paragraph:contains(\"good\")"));
    }

    [Fact]
    public void ExecuteBatch_StopOnErrorStopsBeforeLaterItems()
    {
        var path = CreateBlankDocx();
        const string itemsJson = """
        [
          {"command":"add","parent":"/body/p[99]","type":"paragraph","props":{"text":"bad"}},
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"skipped"}}
        ]
        """;

        using var handler = new WordHandler(path, editable: true);
        var output = BatchExecutor.ExecuteBatch(handler, itemsJson, json: false, stopOnError: true);

        Assert.Contains("[1] ERROR:", output);
        Assert.Contains("Batch complete: 0 succeeded, 1 failed, 1 total", output);
        Assert.Empty(handler.Query("paragraph:contains(\"bad\")"));
        Assert.Empty(handler.Query("paragraph:contains(\"skipped\")"));
    }

    [Fact]
    public void ExecuteBatch_AcceptsPropsArrayAndSkipsMalformedEntries()
    {
        var path = CreateBlankDocx();
        const string itemsJson = """
        [
          {"command":"add","parent":"/body","type":"paragraph","props":["text=Array props","align=center","malformed"]}
        ]
        """;

        using var handler = new WordHandler(path, editable: true);
        var output = BatchExecutor.ExecuteBatch(handler, itemsJson, json: false);
        var paragraph = handler.Get("/body/p[1]");

        Assert.Contains("Batch complete: 1 succeeded, 0 failed, 1 total", output);
        Assert.Equal("Array props", paragraph.Text);
        Assert.Equal("center", Fmt(paragraph)["align"]);
    }

    [Fact]
    public void ExecuteBatch_ObjectPropsCoercesBoolAndNumberValuesToStrings()
    {
        var path = CreateBlankDocx();
        const string itemsJson = """
        [
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Object props","keepNext":true,"spaceBefore":12}}
        ]
        """;

        using var handler = new WordHandler(path, editable: true);
        BatchExecutor.ExecuteBatch(handler, itemsJson, json: false);
        var paragraph = handler.Get("/body/p[1]");

        Assert.Equal("Object props", paragraph.Text);
        Assert.Equal(true, Fmt(paragraph)["keepNext"]);
        Assert.Equal("0.6pt", Fmt(paragraph)["spaceBefore"]);
    }

    [Fact]
    public void ExecuteBatch_MalformedPropsArrayReturnsJsonErrorEnvelope()
    {
        var path = CreateBlankDocx();
        const string itemsJson = """
        [
          {"command":"add","parent":"/body","type":"paragraph","props":[123]}
        ]
        """;

        using var handler = new WordHandler(path, editable: true);
        var output = BatchExecutor.ExecuteBatch(handler, itemsJson, json: true);

        Assert.Contains("\"success\": false", output);
        Assert.Contains("Expected \\\"key=value\\\" string in props array", output);
        Assert.Empty(handler.Query("paragraph"));
    }
}
