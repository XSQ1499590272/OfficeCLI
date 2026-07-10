using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordMutationPositionTests
{
    [Fact]
    public void Resolve_UsesIndexWithoutCallingAnchorFinder()
    {
        var called = false;

        var resolved = InsertPosition.AtIndex(2).Resolve(_ =>
        {
            called = true;
            return 0;
        }, childCount: 4);

        Assert.Equal(2, resolved);
        Assert.False(called);
    }

    [Fact]
    public void Resolve_UsesAnchorModesAndAppendsAfterLastChild()
    {
        Assert.Equal(1, InsertPosition.BeforeElement("/body/p[2]")
            .Resolve(path => path == "/body/p[2]" ? 1 : throw new InvalidOperationException(), 4));
        Assert.Equal(2, InsertPosition.AfterElement("/body/p[2]")
            .Resolve(path => path == "/body/p[2]" ? 1 : throw new InvalidOperationException(), 4));
        Assert.Null(InsertPosition.AfterElement("/body/p[4]")
            .Resolve(_ => 3, 4));
        Assert.Null(new InsertPosition().Resolve(_ => throw new InvalidOperationException(), 4));
    }

    [Fact]
    public void Resolve_PreservesRawIndexBoundaryValuesForCallerValidation()
    {
        Assert.Equal(-1, InsertPosition.AtIndex(-1).Resolve(_ => 0, 3));
        Assert.Equal(99, InsertPosition.AtIndex(99).Resolve(_ => 0, 3));
    }

    [Fact]
    public void Resolve_PropagatesAnchorLookupFailure()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            InsertPosition.BeforeElement("/body/p[99]")
                .Resolve(_ => throw new ArgumentException("path not found"), 2));

        Assert.Contains("path not found", exception.Message);
    }
}
