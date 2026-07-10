using System.IO.Compression;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordStrictAttributeSanitizerTests : WordTestBase
{
    [Fact]
    public void Open_SanitizesInvalidStrictWordAttributes()
    {
        var path = CreateBlankDocx();
        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new()
            {
                ["text"] = "strict attrs",
                ["bold"] = "true",
                ["align"] = "center"
            });
        }
        RewriteDocumentXml(path, xml => xml
            .Replace("<w:b />", "<w:b w:val=\"yes\"/>")
            .Replace("<w:jc w:val=\"center\" />", "<w:jc w:val=\"banana\"/>"));

        using var reopened = new WordHandler(path, editable: true);
        var paragraph = reopened.Get("/body/p[1]");

        Assert.Equal(true, Fmt(paragraph)["bold"]);
        Assert.False(Fmt(paragraph).ContainsKey("align"));
        Assert.Empty(reopened.Validate());
    }

    [Fact]
    public void FailedOpen_ReleasesBackingFileStream()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_corrupt_{Guid.NewGuid():N}.docx");
        File.WriteAllText(path, "not a zip");

        try
        {
            Assert.ThrowsAny<Exception>(() => new WordHandler(path, editable: true));
            using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(stream.CanWrite);
        }
        finally
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }

    private static void RewriteDocumentXml(string path, Func<string, string> rewrite)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = zip.GetEntry("word/document.xml")!;
        string xml;
        using (var reader = new StreamReader(entry.Open()))
            xml = reader.ReadToEnd();
        entry.Delete();
        var replacement = zip.CreateEntry("word/document.xml", CompressionLevel.Optimal);
        using var writer = new StreamWriter(replacement.Open());
        writer.Write(rewrite(xml));
    }
}
