using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordNotesAndCommentsTests : WordTestBase
{
    [Fact]
    public void AddFootnoteAndEndnote_ReadBackTextIdsAndQueryPaths()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Note host" });
        var footnotePath = handler.Add(paragraphPath, "footnote", null, new() { ["text"] = "Footnote text" });
        var endnotePath = handler.Add(paragraphPath, "endnote", null, new() { ["text"] = "Endnote text" });

        var footnote = handler.Get(footnotePath);
        var endnote = handler.Get(endnotePath);

        Assert.Equal("Footnote text", footnote.Text);
        Assert.Equal(1, Convert.ToInt32(Fmt(footnote)["id"]));
        Assert.Equal(footnotePath, Assert.Single(handler.Query("footnote:contains(\"Footnote\")")).Path);
        Assert.Equal("Endnote text", endnote.Text);
        Assert.Equal(1, Convert.ToInt32(Fmt(endnote)["id"]));
        Assert.Equal(endnotePath, Assert.Single(handler.Query("endnote:contains(\"Endnote\")")).Path);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddComment_ReadsBackAuthorInitialsDoneAndAnchor()
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
