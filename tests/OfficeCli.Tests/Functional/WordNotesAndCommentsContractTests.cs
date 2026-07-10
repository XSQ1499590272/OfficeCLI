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
    public void AddNotes_ReadsBackFormattingAndRemoveCleansReferences()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Styled note host" });
        var footnotePath = handler.Add(paragraphPath, "footnote", null, new()
        {
            ["text"] = "Styled footnote",
            ["font"] = "Aptos",
            ["size"] = "11pt",
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "#123456",
            ["underline"] = "single",
            ["strike"] = "true",
            ["highlight"] = "yellow"
        });
        var endnotePath = handler.Add(paragraphPath, "endnote", null, new()
        {
            ["text"] = "Styled endnote",
            ["font"] = "Courier New",
            ["size"] = "10pt",
            ["bold"] = "false",
            ["italic"] = "true",
            ["color"] = "#654321",
            ["underline"] = "double",
            ["highlight"] = "green"
        });

        var footnote = handler.Get(footnotePath);
        var endnote = handler.Get(endnotePath);
        var footnoteFormat = Fmt(footnote);
        var endnoteFormat = Fmt(endnote);

        Assert.Equal("Styled footnote", footnote.Text);
        Assert.Equal("Aptos", footnoteFormat["font"]);
        Assert.Equal("11pt", footnoteFormat["size"]);
        Assert.Equal(true, footnoteFormat["bold"]);
        Assert.Equal(true, footnoteFormat["italic"]);
        Assert.Equal("#123456", footnoteFormat["color"]);
        Assert.Equal("single", footnoteFormat["underline"]);
        Assert.Equal(true, footnoteFormat["strike"]);
        Assert.Equal("yellow", footnoteFormat["highlight"]);
        Assert.Equal("Styled endnote", endnote.Text);
        Assert.Equal("Courier New", endnoteFormat["font"]);
        Assert.Equal("10pt", endnoteFormat["size"]);
        Assert.Equal(false, endnoteFormat["bold"]);
        Assert.Equal(true, endnoteFormat["italic"]);
        Assert.Equal("#654321", endnoteFormat["color"]);
        Assert.Equal("double", endnoteFormat["underline"]);
        Assert.Equal("green", endnoteFormat["highlight"]);

        handler.Remove(footnotePath);
        handler.Remove(endnotePath);

        Assert.Empty(handler.Query("footnote:contains(\"Styled\")"));
        Assert.Empty(handler.Query("endnote:contains(\"Styled\")"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddNotes_DirectionReadbackAndAlignmentRawXmlMatchCurrentSurface()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Direction host" });
        var footnotePath = handler.Add(paragraphPath, "footnote", null, new()
        {
            ["text"] = "RTL footnote",
            ["direction"] = "rtl",
            ["align"] = "right"
        });
        var endnotePath = handler.Add(paragraphPath, "endnote", null, new()
        {
            ["text"] = "RTL endnote",
            ["direction"] = "rtl",
            ["align"] = "center"
        });

        var footnote = handler.Get(footnotePath);
        var endnote = handler.Get(endnotePath);
        Assert.Equal("rtl", Fmt(footnote)["direction"]);
        Assert.False(Fmt(footnote).ContainsKey("align"));
        Assert.Equal("rtl", Fmt(endnote)["direction"]);
        Assert.False(Fmt(endnote).ContainsKey("align"));

        var footnotesRaw = handler.Raw("/footnotes");
        var endnotesRaw = handler.Raw("/endnotes");
        Assert.Contains("w:bidi", footnotesRaw);
        Assert.Contains("w:val=\"right\"", footnotesRaw);
        Assert.Contains("w:bidi", endnotesRaw);
        Assert.Contains("w:val=\"center\"", endnotesRaw);
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

    [Fact]
    public void AddComment_ReadsBackExtendedMetadataReplyThreadAndFilters()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var parentParagraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Parent anchor" });
        var replyParagraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Reply anchor" });
        var parentCommentPath = handler.Add(parentParagraphPath, "comment", null, new()
        {
            ["text"] = "Resolve this",
            ["author"] = "Reviewer One",
            ["initials"] = "R1",
            ["date"] = "2026-07-08T01:02:03Z",
            ["done"] = "true",
            ["direction"] = "rtl"
        });
        var replyCommentPath = handler.Add(replyParagraphPath, "comment", null, new()
        {
            ["text"] = "Reply thread",
            ["author"] = "Reviewer Two",
            ["initials"] = "R2",
            ["date"] = "2026-07-08T04:05:06Z",
            ["parentId"] = "1",
            ["resolved"] = "false"
        });

        var parent = handler.Get(parentCommentPath);
        var reply = handler.Get(replyCommentPath);
        var parentFormat = Fmt(parent);
        var replyFormat = Fmt(reply);

        Assert.Equal("Resolve this", parent.Text);
        Assert.Equal("Reviewer One", parentFormat["author"]);
        Assert.Equal("R1", parentFormat["initials"]);
        Assert.Equal("2026-07-08T01:02:03.0000000Z", parentFormat["date"]);
        Assert.Equal("true", parentFormat["done"]);
        Assert.Equal("rtl", parentFormat["direction"]);
        Assert.Equal(parentParagraphPath, parentFormat["anchoredTo"]);
        Assert.Equal("Reply thread", reply.Text);
        Assert.Equal("Reviewer Two", replyFormat["author"]);
        Assert.Equal("R2", replyFormat["initials"]);
        Assert.Equal("2026-07-08T04:05:06.0000000Z", replyFormat["date"]);
        Assert.Equal("false", replyFormat["done"]);
        Assert.Equal("1", replyFormat["parentId"]);
        Assert.Equal(replyParagraphPath, replyFormat["anchoredTo"]);
        Assert.Contains(handler.Query("comment[done=true]"), node => node.Text == "Resolve this");
        Assert.Contains(handler.Query("comment[parentId=1]"), node => node.Text == "Reply thread");
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void RemoveComment_CleansRangeMarkersAndCommentDefinition()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var startParagraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Range start" });
        var endParagraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Range end" });
        var commentPath = handler.Add(startParagraphPath, "comment", null, new()
        {
            ["text"] = "Span comment",
            ["rangeOpen"] = "true"
        });
        handler.Add(endParagraphPath, "comment", null, new()
        {
            ["rangeEnd"] = "true"
        });

        var comment = handler.Get(commentPath);
        Assert.Equal(startParagraphPath, Fmt(comment)["anchoredTo"]);

        handler.Remove(commentPath);

        Assert.Empty(handler.Query("comment:contains(\"Span\")"));
        Assert.Empty(handler.Validate());
    }
}
