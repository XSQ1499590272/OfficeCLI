using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordDumpFilterTests : WordTestBase
{
    [Fact]
    public void DumpFiltersDerivedAndUnstableParagraphProperties()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/", new() { ["docDefaults.font"] = "Georgia" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "dump filter" });

        var (items, warnings) = WordBatchEmitter.EmitWordWithWarnings(handler, "/body");
        var paragraph = Assert.Single(items, item => item.Command == "add" && item.Type == "p");
        var props = paragraph.Props!;

        Assert.DoesNotContain("paraId", props.Keys, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain("textId", props.Keys, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain("relId", props.Keys, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain(props.Keys, key => key.StartsWith("effective.", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain("styleId", props.Keys, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain("styleName", props.Keys, StringComparer.OrdinalIgnoreCase);
        Assert.Empty(warnings);
    }
}
