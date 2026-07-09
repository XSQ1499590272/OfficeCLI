using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordMoveCopyContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void Move_ReordersBodyParagraphsWithoutLosingText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var first = handler.Add("/body", "paragraph", null, new() { ["text"] = "first" });
        var second = handler.Add("/body", "paragraph", null, new() { ["text"] = "second" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "third" });

        handler.Move(first, "/body", InsertPosition.AfterElement(second));
        var body = handler.Get("/body", depth: 1);
        var paragraphs = body.Children.Where(child => child.Type == "paragraph").Select(child => child.Text).ToList();

        Assert.Equal(new[] { "second", "first", "third" }, paragraphs);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void CopyFrom_ClonesParagraphIntoBody()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var source = handler.Add("/body", "paragraph", null, new() { ["text"] = "copy me" });

        var copyPath = handler.CopyFrom(source, "/body", null);
        var paragraphs = handler.Query("paragraph:contains(\"copy me\")");

        Assert.Equal("copy me", handler.Get(copyPath).Text);
        Assert.Equal(2, paragraphs.Count);
        Assert.Empty(handler.Validate());
    }
}
