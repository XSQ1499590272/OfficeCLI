using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordPathAliasTests
{
    [Theory]
    [InlineData("paragraph", "p")]
    [InlineData("PARAGRAPH", "p")]
    [InlineData("run", "r")]
    [InlineData("table", "tbl")]
    [InlineData("row", "tr")]
    [InlineData("cell", "tc")]
    [InlineData("picture", "picture")]
    public void Resolve_MapsWordAliasesCaseInsensitively(string input, string expected)
    {
        Assert.Equal(expected, PathAliases.Resolve(input));
    }

    [Theory]
    [InlineData("body", "body")]
    [InlineData("customPart", "customPart")]
    [InlineData("tblPr", "tblPr")]
    public void Resolve_LeavesUnknownSegmentsUntouched(string input, string expected)
    {
        Assert.Equal(expected, PathAliases.Resolve(input));
    }
}
