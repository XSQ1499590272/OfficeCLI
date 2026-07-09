using System.Reflection;
using OfficeCli;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class PathIndexTests
{
    [Theory]
    [InlineData(1, 0)]
    [InlineData(2, 1)]
    [InlineData(0, -1)]
    public void ToArrayIndex_UsesCurrentOneBasedToZeroBasedArithmetic(int xpathIndex, int expected)
    {
        Assert.Equal(expected, InvokePathIndex("ToArrayIndex", xpathIndex));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 2)]
    [InlineData(-1, 0)]
    public void FromArrayIndex_UsesCurrentZeroBasedToOneBasedArithmetic(int arrayIndex, int expected)
    {
        Assert.Equal(expected, InvokePathIndex("FromArrayIndex", arrayIndex));
    }

    private static int InvokePathIndex(string methodName, int value)
    {
        var type = typeof(BlankDocCreator).Assembly.GetType("OfficeCli.Core.PathIndex")
            ?? throw new InvalidOperationException("PathIndex type was not found.");
        var method = type.GetMethod(methodName, BindingFlags.Public | BindingFlags.Static)
            ?? throw new InvalidOperationException($"{methodName} was not found.");
        return (int)method.Invoke(null, [value])!;
    }
}
