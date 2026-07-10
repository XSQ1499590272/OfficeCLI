using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordNumberingAndLocaleTests
{
    [Theory]
    [InlineData(0, "decimal", "0")]
    [InlineData(-3, "decimal", "-3")]
    [InlineData(3, "decimalzero", "03")]
    [InlineData(4, "upperroman", "IV")]
    [InlineData(27, "lowerletter", "aa")]
    [InlineData(22, "ordinal", "22nd")]
    [InlineData(21, "cardinaltext", "Twenty-one")]
    [InlineData(12, "chinesecounting", "十二")]
    [InlineData(16, "hex", "10")]
    [InlineData(1, "bullet", "•")]
    [InlineData(1, "none", "")]
    [InlineData(17, "future-format", "17")]
    public void Render_ReturnsStableRepresentativeNumberFormats(int number, string format, string expected)
    {
        Assert.Equal(expected, WordNumFmtRenderer.Render(number, format));
    }

    [Fact]
    public void Render_ClampsNonDecimalFormatsAtOne()
    {
        Assert.Equal("I", WordNumFmtRenderer.Render(0, "upperroman"));
        Assert.Equal("a", WordNumFmtRenderer.Render(-1, "lowerletter"));
        Assert.Equal("01", WordNumFmtRenderer.Render(0, "decimalzero"));
    }

    [Fact]
    public void Render_FallsBackToDecimalOutsideRomanRange()
    {
        Assert.Equal("4000", WordNumFmtRenderer.Render(4000, "upperroman"));
    }

    [Theory]
    [InlineData(1, "decimalenclosedcircle", "①")]
    [InlineData(20, "decimalenclosedcircle", "⑳")]
    [InlineData(21, "decimalenclosedcircle", "21")]
    [InlineData(3, "decimalenclosedfullstop", "⒊")]
    [InlineData(4, "decimalenclosedparen", "⑷")]
    [InlineData(3, "decimalfullwidth", "３")]
    [InlineData(1, "arabicabjad", "أ")]
    [InlineData(2, "arabicalpha", "ب")]
    [InlineData(16, "hebrew1", "טז")]
    [InlineData(10, "thainumbers", "๑๐")]
    [InlineData(25, "hindinumbers", "२५")]
    [InlineData(1, "aiueo", "ア")]
    [InlineData(46, "aiueo", "ン")]
    [InlineData(1, "chosung", "ㄱ")]
    [InlineData(14, "ganada", "하")]
    [InlineData(1, "ideographtraditional", "甲")]
    [InlineData(12, "ideographzodiac", "亥")]
    [InlineData(60, "ideographzodiactraditional", "癸亥")]
    [InlineData(5, "chicago", "**")]
    [InlineData(3, "ideographenclosedcircle", "㈢")]
    [InlineData(7, "numberindash", "- 7 -")]
    [InlineData(101, "chinesecountingthousand", "一百〇一")]
    [InlineData(2025, "japanesecounting", "二千二十五")]
    [InlineData(25, "koreancounting", "이십오")]
    [InlineData(25, "koreanlegal", "스물다섯")]
    [InlineData(4000, "hex", "FA0")]
    public void Render_CoversExtendedEcmaNumberFormats(int number, string format, string expected)
    {
        Assert.Equal(expected, WordNumFmtRenderer.Render(number, format));
    }

    [Fact]
    public void Resolve_UsesRegionalAndScriptLocaleDefaults()
    {
        Assert.Equal(("Times New Roman", "等线", (string?)null), LocaleFontRegistry.Resolve("zh_CN"));
        Assert.Equal(("Times New Roman", "新細明體", (string?)null), LocaleFontRegistry.Resolve("zh-TW"));
        Assert.Equal(("Times New Roman", (string?)null, "Arabic Typesetting"), LocaleFontRegistry.Resolve("ar-SA"));
        Assert.Equal(((string?)null, (string?)null, (string?)null), LocaleFontRegistry.Resolve(null));
    }

    [Fact]
    public void LocaleHelpers_ExposeDirectionAndCjkFallbacks()
    {
        Assert.True(LocaleFontRegistry.IsRightToLeft("ar-SA"));
        Assert.True(LocaleFontRegistry.IsRightToLeft("he"));
        Assert.False(LocaleFontRegistry.IsRightToLeft("en-US"));
        Assert.Contains("PingFang SC", LocaleFontRegistry.GetCjkCssFallback("zh-CN"));
        Assert.Contains("Hiragino Sans", LocaleFontRegistry.GetCjkCssFallback("ja"));
        Assert.Equal("ja", LocaleFontRegistry.DetectLocaleFromCjkFontName("Yu Mincho"));
        Assert.Equal("zh", LocaleFontRegistry.DetectLocaleFromCjkFontName("Microsoft YaHei"));
        Assert.Null(LocaleFontRegistry.DetectLocaleFromCjkFontName("Calibri"));
    }

    [Fact]
    public void PageDefaults_ValidateEcmaDimensionBoundaries()
    {
        WordPageDefaults.ValidatePageDim(WordPageDefaults.PageDimMinTwips, "width");
        WordPageDefaults.ValidatePageDim(WordPageDefaults.PageDimMaxTwips, "height");

        var tooSmall = Assert.Throws<ArgumentException>(() =>
            WordPageDefaults.ValidatePageDim(WordPageDefaults.PageDimMinTwips - 1, "width"));
        var tooLarge = Assert.Throws<ArgumentException>(() =>
            WordPageDefaults.ValidatePageDim(WordPageDefaults.PageDimMaxTwips + 1L, "height"));

        Assert.Contains("width", tooSmall.Message);
        Assert.Contains("height", tooLarge.Message);
    }

    [Fact]
    public void ResolveEffectiveLocale_PrefersExplicitAndIgnoresShellSentinels()
    {
        var previous = LocaleFontRegistry.OsLocaleSnapshot;
        try
        {
            LocaleFontRegistry.OsLocaleSnapshot = "zh-CN";
            Assert.Equal("zh-CN", LocaleFontRegistry.ResolveEffectiveLocale(null));
            Assert.Equal("ja-JP", LocaleFontRegistry.ResolveEffectiveLocale("ja-JP"));

            LocaleFontRegistry.OsLocaleSnapshot = "C";
            Assert.Null(LocaleFontRegistry.ResolveEffectiveLocale(null));
            LocaleFontRegistry.OsLocaleSnapshot = "POSIX";
            Assert.Null(LocaleFontRegistry.ResolveEffectiveLocale(null));
        }
        finally
        {
            LocaleFontRegistry.OsLocaleSnapshot = previous;
        }
    }
}
