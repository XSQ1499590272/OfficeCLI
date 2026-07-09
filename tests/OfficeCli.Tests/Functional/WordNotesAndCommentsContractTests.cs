using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordNotesAndCommentsContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddFootnote_ReadsBackTextIdAndValidates()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Footnote host" });
        var footnotePath = handler.Add(paragraphPath, "footnote", null, new() { ["text"] = "Footnote text" });

        var footnote = handler.Get(footnotePath);
        var queried = Assert.Single(handler.Query("footnote:contains(\"Footnote\")"));

        Assert.Equal("footnote", footnote.Type);
        Assert.Equal("Footnote text", footnote.Text);
        Assert.Equal(1, Convert.ToInt32(Fmt(footnote)["id"]));
        Assert.Equal(footnotePath, queried.Path);
        Assert.Equal("Footnote text", queried.Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddEndnote_ReadsBackTextIdAndValidates()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Endnote host" });
        var endnotePath = handler.Add(paragraphPath, "endnote", null, new() { ["text"] = "Endnote text" });

        var endnote = handler.Get(endnotePath);
        var queried = Assert.Single(handler.Query("endnote:contains(\"Endnote\")"));

        Assert.Equal("endnote", endnote.Type);
        Assert.Equal("Endnote text", endnote.Text);
        Assert.Equal(1, Convert.ToInt32(Fmt(endnote)["id"]));
        Assert.Equal(endnotePath, queried.Path);
        Assert.Equal("Endnote text", queried.Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddComment_ReadsBackTextAuthorInitialsAndAnchor()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Comment host" });
        var commentPath = handler.Add(paragraphPath, "comment", null, new()
        {
            ["text"] = "Review this",
            ["author"] = "Office Agent",
            ["initials"] = "OA",
            ["date"] = "2026-07-08T00:00:00Z"
        });

        var comment = handler.Get(commentPath);
        var queried = Assert.Single(handler.Query("comment:contains(\"Review\")"));

        Assert.Equal("comment", comment.Type);
        Assert.Equal("Review this", comment.Text);
        Assert.Equal("Office Agent", Fmt(comment)["author"]);
        Assert.Equal("OA", Fmt(comment)["initials"]);
        Assert.Equal("1", Fmt(comment)["id"]);
        Assert.Equal("false", Fmt(comment)["done"]);
        Assert.True(Fmt(comment).ContainsKey("anchoredTo"));
        Assert.Equal("/comments/comment[@commentId=1]", queried.Path);
        Assert.Equal("Office Agent", Fmt(queried)["author"]);
        Assert.Empty(handler.Validate());
    }
}
