using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public class WordComplexRelationshipCleanupTests : OfficeCli.Tests.Unit.WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void RemoveRelationshipObjectsLeavesNoDanglingPartsOrMarkers()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "remove-picture.png"
        });
        var picturePath = Assert.Single(handler.Query("picture")).Path!;
        handler.Remove(picturePath);

        handler.Add("/body", "ole", null, new()
        {
            ["src"] = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,SGVsbG8=",
            ["progId"] = "Word.Document.12",
            ["contentType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ["name"] = "remove-ole"
        });
        var olePath = Assert.Single(handler.Query("ole")).Path!;
        handler.Remove(olePath);

        handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "Q1,Q2",
            ["data"] = "Revenue:10,20"
        });
        var chartPath = Assert.Single(handler.Query("chart")).Path!;
        handler.Remove(chartPath);

        var hyperlinkHost = handler.Add("/body", "paragraph", null, new() { ["text"] = "Link host" });
        var hyperlinkPath = handler.Add(hyperlinkHost, "hyperlink", null, new()
        {
            ["url"] = "https://example.com/remove",
            ["text"] = "remove-link"
        });
        handler.Remove(hyperlinkPath);

        var notesHost = handler.Add("/body", "paragraph", null, new() { ["text"] = "Notes host" });
        var footnotePath = handler.Add(notesHost, "footnote", null, new() { ["text"] = "remove-footnote" });
        var endnotePath = handler.Add(notesHost, "endnote", null, new() { ["text"] = "remove-endnote" });
        var commentPath = handler.Add(notesHost, "comment", null, new() { ["text"] = "remove-comment" });
        handler.Remove(commentPath);
        handler.Remove(footnotePath);
        handler.Remove(endnotePath);

        var headerPath = handler.Add("/", "header", null, new() { ["text"] = "remove-header" });
        var footerPath = handler.Add("/", "footer", null, new() { ["text"] = "remove-footer" });
        handler.Remove(headerPath);
        handler.Remove(footerPath);

        Assert.Empty(handler.Query("picture"));
        Assert.Empty(handler.Query("ole"));
        Assert.Empty(handler.Query("chart"));
        Assert.Empty(handler.Query("hyperlink"));
        Assert.Empty(handler.Query("footnote"));
        Assert.Empty(handler.Query("endnote"));
        Assert.Empty(handler.Query("comment"));
        Assert.Empty(handler.Query("header"));
        Assert.Empty(handler.Query("footer"));
        Assert.Empty(handler.Validate());
    }
}
