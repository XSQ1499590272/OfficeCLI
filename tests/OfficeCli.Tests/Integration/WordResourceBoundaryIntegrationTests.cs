using System.Text.Json;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordResourceBoundaryIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void LongTextThroughCli_RoundTripsWithoutTruncation()
    {
        var path = CreateBlankDocx();
        var text = new string('x', 128 * 1024);

        Assert.Equal(0, Invoke("add", path, "/body", "--type", "paragraph", "--prop", $"text={text}"));

        using var handler = new WordHandler(path, editable: false);
        Assert.Equal(text, handler.Query("paragraph").Single().Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void GetWithExcessiveDepth_IsCappedAndDoesNotFail()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
            setup.Add("/body", "paragraph", null, new() { ["text"] = "depth boundary" });

        Assert.Equal(0, Invoke("get", path, "/body/p[1]", "--depth", "100000", "--json"));
    }

    [Fact]
    public void LargeBatch_CompletesAndPersistsEveryItem()
    {
        var path = CreateBlankDocx();
        var items = Enumerable.Range(0, 512)
            .Select(index => new
            {
                command = "add",
                parent = "/body",
                type = "paragraph",
                props = new Dictionary<string, string> { ["text"] = $"batch-{index}" }
            })
            .ToArray();
        var batchJson = JsonSerializer.Serialize(items);

        using (var handler = new WordHandler(path, editable: true))
        {
            var output = BatchExecutor.ExecuteBatch(handler, batchJson, json: true);
            using var result = JsonDocument.Parse(output);
            Assert.True(result.RootElement.GetProperty("success").GetBoolean(), output);
            handler.Save();
        }

        using var reopened = new WordHandler(path, editable: false);
        var texts = reopened.Query("paragraph").Select(node => node.Text).ToList();
        Assert.Equal(512, texts.Count);
        Assert.Contains("batch-0", texts);
        Assert.Contains("batch-511", texts);
        Assert.Empty(reopened.Validate());
    }

    private static int Invoke(params string[] args)
    {
        var root = CommandBuilder.BuildRootCommand();
        return root.Parse(args).Invoke();
    }
}
