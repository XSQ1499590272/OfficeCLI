using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordErrorContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void FailedAdd_InvalidParentPath_DoesNotPersistPartialParagraph()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/body/p[99]", "paragraph", null, new() { ["text"] = "should not appear" }));
        }

        using var reopened = new WordHandler(path, editable: false);
        Assert.DoesNotContain(reopened.Query("paragraph"), node => node.Text == "should not appear");
    }

    [Fact]
    public void FailedAdd_PictureWithoutSrc_DoesNotPersistPartialPicture()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/body", "picture", null, new() { ["alt"] = "missing source" }));
        }

        using var reopened = new WordHandler(path, editable: false);
        Assert.Empty(reopened.Query("picture"));
    }

    [Fact]
    public void FailedAdd_PictureWithZeroWidth_DoesNotPersistPartialPicture()
    {
        const string tinyPngDataUri =
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/body", "picture", null, new()
                {
                    ["src"] = tinyPngDataUri,
                    ["width"] = "0"
                }));
        }

        using var reopened = new WordHandler(path, editable: false);
        Assert.Empty(reopened.Query("picture"));
    }

    [Fact]
    public void FailedSet_RowInvalidAlign_DoesNotPersistPartialRowProperties()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

            Assert.Throws<ArgumentException>(() =>
                handler.Set("/body/tbl[1]/tr[1]", new() { ["rowAlign"] = "diagonal" }));
        }

        using var reopened = new WordHandler(path, editable: false);
        var row = reopened.Get("/body/tbl[1]/tr[1]");
        Assert.False(Fmt(row).ContainsKey("rowAlign"));
    }

    [Fact]
    public void SetUnsupportedParagraphProperty_ReturnsUnsupportedAndDoesNotApplyProperty()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "stable" });

        var unsupported = handler.Set(paragraphPath, new() { ["madeUpProperty"] = "value" });
        var paragraph = handler.Get(paragraphPath);

        Assert.Contains(unsupported, item => item.Contains("madeUpProperty"));
        Assert.Equal("stable", paragraph.Text);
        Assert.False(Fmt(paragraph).ContainsKey("madeUpProperty"));
    }

    [Fact]
    public void OpenCorruptDocx_ThrowsAndLeavesOriginalBytesUntouched()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_corrupt_{Guid.NewGuid():N}.docx");
        var original = new byte[] { 0x6E, 0x6F, 0x74, 0x2D, 0x64, 0x6F, 0x63, 0x78 };
        File.WriteAllBytes(path, original);

        try
        {
            Assert.ThrowsAny<Exception>(() => new WordHandler(path, editable: true));
            Assert.Equal(original, File.ReadAllBytes(path));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }
}
