using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordResourceLimitTests
{
    [Fact]
    public void ResourceLimits_ExposeStableDocumentAndRegexBounds()
    {
        Assert.Equal(256, DocumentLimits.MaxRecursionDepth);
        Assert.Equal(100_000, DocumentLimits.MaxZipEntries);
        Assert.Equal(1000, DocumentLimits.MaxCompressionRatio);
        Assert.Equal(5, DocumentLimits.RegexMatchTimeout.TotalSeconds);
        Assert.Equal(2L * 1024 * 1024 * 1024, DocumentLimits.MaxUncompressedBytes);
    }

    [Fact]
    public void EnsureDepth_AllowsConfiguredBoundary()
    {
        var exception = Record.Exception(() => DocumentLimits.EnsureDepth(DocumentLimits.MaxRecursionDepth));

        Assert.Null(exception);
    }

    [Fact]
    public void EnsureDepth_RejectsDepthBeyondConfiguredBoundaryWithStructuredCode()
    {
        var exception = Assert.Throws<CliException>(() =>
            DocumentLimits.EnsureDepth(DocumentLimits.MaxRecursionDepth + 1));

        Assert.Equal("max_depth_exceeded", exception.Code);
        Assert.Contains("maximum supported depth", exception.Message);
        Assert.NotNull(exception.Suggestion);
    }
}
