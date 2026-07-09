using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordErrorBoundaryTests : WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void FailedAdd_InvalidParentAndPictureProps_DoNotPersistPartialNodes()
    {
        var path = CreateBlankDocx();

        using (var handler = new WordHandler(path, editable: true))
        {
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/body/p[99]", "paragraph", null, new() { ["text"] = "should not appear" }));
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/body", "picture", null, new() { ["alt"] = "missing source" }));
            Assert.Throws<ArgumentException>(() =>
                handler.Add("/body", "picture", null, new()
                {
                    ["src"] = TinyPngDataUri,
                    ["width"] = "0"
                }));
        }

        using var reopened = new WordHandler(path, editable: false);
        Assert.DoesNotContain(reopened.Query("paragraph"), node => node.Text == "should not appear");
        Assert.Empty(reopened.Query("picture"));
    }

    [Fact]
    public void FailedSet_InvalidOrUnsupportedProperties_DoNotApplyValues()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "stable" });
        handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        var unsupported = handler.Set(paragraphPath, new() { ["madeUpProperty"] = "value" });
        Assert.Throws<ArgumentException>(() =>
            handler.Set("/body/tbl[1]/tr[1]", new() { ["rowAlign"] = "diagonal" }));

        var paragraph = handler.Get(paragraphPath);
        var row = handler.Get("/body/tbl[1]/tr[1]");

        Assert.Contains(unsupported, item => item.Contains("madeUpProperty"));
        Assert.Equal("stable", paragraph.Text);
        Assert.False(Fmt(paragraph).ContainsKey("madeUpProperty"));
        Assert.False(Fmt(row).ContainsKey("rowAlign"));
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
