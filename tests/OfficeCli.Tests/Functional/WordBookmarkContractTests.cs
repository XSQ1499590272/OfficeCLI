using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordBookmarkContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddBookmark_ReadsBackNameAndQueryableBookmark()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "bookmark host" });
        var bookmarkPath = handler.Add(paragraphPath, "bookmark", null, new() { ["name"] = "TestMark" });

        var bookmark = handler.Get(bookmarkPath);
        var queried = Assert.Single(handler.Query("bookmark"));

        Assert.Equal("bookmark", bookmark.Type);
        Assert.Equal("TestMark", Fmt(bookmark)["name"]);
        Assert.Equal("/bookmark[@name=TestMark]", queried.Path);
        Assert.Equal("TestMark", Fmt(queried)["name"]);
    }

    [Fact]
    public void FailedAdd_BookmarkWithoutName_DoesNotPersistBookmark()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "bookmark host" });
            Assert.Throws<ArgumentException>(() =>
                handler.Add(paragraphPath, "bookmark", null, new()));
        }

        using var reopened = new WordHandler(path, editable: false);
        Assert.Empty(reopened.Query("bookmark"));
    }

    [Fact]
    public void AddBookmark_AllowsDuplicateAndPathSpecialNames()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "bookmark host" });

        handler.Add(paragraphPath, "bookmark", null, new() { ["name"] = "DuplicateMark" });
        handler.Add(paragraphPath, "bookmark", null, new() { ["name"] = "DuplicateMark" });
        handler.Add(paragraphPath, "bookmark", null, new() { ["name"] = "Review/Analysis" });

        var bookmarks = handler.Query("bookmark");

        Assert.Equal(2, bookmarks.Count(node => Equals(Fmt(node)["name"], "DuplicateMark")));
        Assert.Contains(bookmarks, node => Equals(Fmt(node)["name"], "Review/Analysis"));
        Assert.Empty(handler.Validate());
    }
}
