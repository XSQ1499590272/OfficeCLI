using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class WordHtmlRefreshTests : WordTestBase
{
    [Theory]
    [InlineData(" PAGEREF _Toc123 \\h ", "_Toc123")]
    [InlineData("PAGEREF _Toc9", "_Toc9")]
    [InlineData(" REF _Toc123 \\h ", null)]
    [InlineData("", null)]
    public void ExtractPagerefAnchor_UsesOnlyPagerefInstruction(
        string instruction,
        string? expected)
    {
        var method = typeof(WordHtmlRefresh).GetMethod(
            "ExtractPagerefAnchor",
            BindingFlags.Static | BindingFlags.NonPublic);

        Assert.NotNull(method);
        Assert.Equal(expected, method!.Invoke(null, [instruction]));
    }

    [Fact]
    public void ApplyPageNumbers_RewritesCachedPagerefResultForKnownAnchor()
    {
        var path = CreateBlankDocx();

        using var document = WordprocessingDocument.Open(path, true);
        var body = document.MainDocumentPart!.Document!.Body!;
        var paragraph = new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode { Text = " PAGEREF _Toc123 \\h " }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("0")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        body.InsertBefore(paragraph, body.GetFirstChild<SectionProperties>());

        var method = typeof(WordHtmlRefresh).GetMethod(
            "ApplyPageNumbers",
            BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(method);
        method!.Invoke(null, [document, new Dictionary<string, int> { ["_Toc123"] = 7 }]);

        Assert.Equal("7", paragraph.Descendants<Text>().Last().Text);
        document.MainDocumentPart.Document.Save();
    }

    [Fact]
    public void RefreshViaHtml_MissingDocumentReturnsFalse()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_missing_{Guid.NewGuid():N}.docx");

        Assert.False(WordHtmlRefresh.RefreshViaHtml(path));
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void RefreshViaHtml_CorruptDocumentReturnsFalseWithoutChangingBytes()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_refresh_corrupt_{Guid.NewGuid():N}.docx");
        var original = new byte[] { 0x6E, 0x6F, 0x74, 0x2D, 0x64, 0x6F, 0x63, 0x78 };
        File.WriteAllBytes(path, original);

        try
        {
            Assert.False(WordHtmlRefresh.RefreshViaHtml(path));
            Assert.Equal(original, File.ReadAllBytes(path));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }
}
