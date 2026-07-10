using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordRuntimePolicyTests
{
    [Fact]
    public void TryParse_RecognizesNamedAndFixedFlushModes()
    {
        AssertPolicy("each", ResidentFlushMode.Each, null);
        AssertPolicy(" AUTO ", ResidentFlushMode.Auto, null);
        AssertPolicy("off", ResidentFlushMode.Off, null);
        AssertPolicy("0", ResidentFlushMode.Off, null);
        AssertPolicy("-1", ResidentFlushMode.Off, null);
        AssertPolicy("3", ResidentFlushMode.Fixed, TimeSpan.FromSeconds(3));
    }

    [Fact]
    public void TryParse_RejectsBlankAndOutOfRangeValues()
    {
        Assert.False(ResidentFlushPolicy.TryParse(null, 2, 10, out _, out _));
        Assert.False(ResidentFlushPolicy.TryParse("   ", 2, 10, out _, out _));
        Assert.False(ResidentFlushPolicy.TryParse("1", 2, 10, out _, out _));
        Assert.False(ResidentFlushPolicy.TryParse("11", 2, 10, out _, out _));
        Assert.False(ResidentFlushPolicy.TryParse("seconds", 2, 10, out _, out _));
    }

    [Fact]
    public void NextEma_SeedsAndUsesAsymmetricRiseAndFallSmoothing()
    {
        Assert.Equal(2.5, ResidentFlushPolicy.NextEmaSeconds(-1, 2.5));
        Assert.Equal(0, ResidentFlushPolicy.NextEmaSeconds(-1, -3));
        Assert.Equal(3.1, ResidentFlushPolicy.NextEmaSeconds(1, 4), 10);
        Assert.Equal(3.4, ResidentFlushPolicy.NextEmaSeconds(4, 1), 10);
    }

    [Fact]
    public void IntervalForEma_StaysWithinAdaptiveBounds()
    {
        Assert.Equal(2, ResidentFlushPolicy.IntervalForEma(-1).TotalSeconds);
        Assert.Equal(2, ResidentFlushPolicy.IntervalForEma(0.5).TotalSeconds);
        Assert.Equal(6, ResidentFlushPolicy.IntervalForEma(1.5).TotalSeconds);
        Assert.Equal(10, ResidentFlushPolicy.IntervalForEma(3).TotalSeconds);
    }

    private static void AssertPolicy(string raw, ResidentFlushMode expectedMode, TimeSpan? expectedInterval)
    {
        Assert.True(ResidentFlushPolicy.TryParse(raw, 2, 10, out var mode, out var interval));
        Assert.Equal(expectedMode, mode);
        Assert.Equal(expectedInterval ?? default, interval);
    }
}
