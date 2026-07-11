using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class WordDumpResourceOrderTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void DumpFullDocument_EmitsResourcesBeforeBodyInStableDependencyOrder()
    {
        var path = CreateBlankDocx();
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "resource order" });
        var items = WordBatchEmitter.EmitWord(handler);

        var bodyIndex = items.FindIndex(item => item.Command == "add" && item.Parent == "/body");
        Assert.True(bodyIndex > 0);

        var beforeBody = items.Take(bodyIndex).ToList();
        Assert.Contains(beforeBody, item => item.Command == "add" && item.Parent == "/styles");
        Assert.Contains(beforeBody, item => item.Command == "raw-set" && item.Part == "/theme");
        Assert.Contains(beforeBody, item => item.Command == "raw-set" && item.Part == "/settings");

        var coreIndex = IndexOfPart(beforeBody, "/docProps/core.xml");
        var appIndex = IndexOfPart(beforeBody, "/docProps/app.xml");
        var customIndex = IndexOfPart(beforeBody, "/docProps/custom.xml");
        Assert.True(coreIndex >= 0 && coreIndex < appIndex && appIndex < customIndex);
        Assert.DoesNotContain(items.Take(bodyIndex), item => item.Parent == "/body");

        Assert.Equal("p", items[bodyIndex].Type);
        Assert.Equal("resource order", items[bodyIndex].Props!["text"]);
    }

    private static int IndexOfPart(IReadOnlyList<BatchItem> items, string part)
        => items
            .Select((item, index) => (item, index))
            .FirstOrDefault(pair => pair.item.Command == "raw-set" && pair.item.Part == part)
            .index;
}
