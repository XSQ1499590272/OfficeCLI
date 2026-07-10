using DocumentFormat.OpenXml.Packaging;
using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordHtmlSafetyTests
{
    [Fact]
    public void HtmlEncode_EscapesTextAndAttributeDelimiters()
    {
        Assert.Equal("&lt;tag a=&quot;1&quot; b=&#39;2&#39;&gt;&amp;", HtmlPreviewHelper.HtmlEncode("<tag a=\"1\" b='2'>&"));
    }

    [Fact]
    public void HtmlEncode_PreservesUnicodeAndDoesNotDoubleEncodeOrderSensitiveInput()
    {
        Assert.Equal("中文 &amp; العربية &lt;ok&gt;", HtmlPreviewHelper.HtmlEncode("中文 & العربية <ok>"));
        Assert.Equal("&amp;lt;", HtmlPreviewHelper.HtmlEncode("&lt;"));
    }

    [Theory]
    [InlineData("image/png", false)]
    [InlineData("image/jpeg", false)]
    [InlineData("image/wmf", true)]
    [InlineData("image/emf", true)]
    [InlineData("image/tiff", true)]
    public void PartToDataUri_UsesPlaceholderOnlyForUndecodableImageTypes(string contentType, bool placeholder)
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_html_{Guid.NewGuid():N}.docx");
        try
        {
            OfficeCli.BlankDocCreator.Create(path);
            using var doc = WordprocessingDocument.Open(path, true);
            var image = doc.MainDocumentPart!.AddImagePart(contentType switch
            {
                "image/jpeg" => DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg,
                "image/wmf" => DocumentFormat.OpenXml.Packaging.ImagePartType.Wmf,
                "image/emf" => DocumentFormat.OpenXml.Packaging.ImagePartType.Emf,
                "image/tiff" => DocumentFormat.OpenXml.Packaging.ImagePartType.Tiff,
                _ => DocumentFormat.OpenXml.Packaging.ImagePartType.Png
            });
            using (var stream = image.GetStream(FileMode.Create, FileAccess.Write))
                stream.WriteByte(1);

            var mainPart = doc.MainDocumentPart!;
            var result = HtmlPreviewHelper.PartToDataUri(mainPart, mainPart.GetIdOfPart(image));

            Assert.NotNull(result);
            Assert.Equal(placeholder, result!.StartsWith("data:image/svg+xml;base64,", StringComparison.Ordinal));
            if (!placeholder)
                Assert.StartsWith($"data:{contentType};base64,", result);
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }
}
