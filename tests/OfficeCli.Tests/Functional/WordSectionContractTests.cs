using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordSectionContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddSection_ReadsBackLayoutMarginsAndColumns()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var sectionPath = handler.Add("/body", "section", null, new()
        {
            ["type"] = "continuous",
            ["orientation"] = "landscape",
            ["marginTop"] = "1cm",
            ["marginBottom"] = "1cm",
            ["marginLeft"] = "2cm",
            ["marginRight"] = "2cm",
            ["columns"] = "2",
            ["columnSpace"] = "1cm"
        });

        var section = handler.Get(sectionPath);

        Assert.Equal("section", section.Type);
        Assert.Equal("continuous", Fmt(section)["type"]);
        Assert.Equal("landscape", Fmt(section)["orientation"]);
        Assert.Equal("1cm", Fmt(section)["marginTop"]);
        Assert.Equal("1cm", Fmt(section)["marginBottom"]);
        Assert.Equal("2cm", Fmt(section)["marginLeft"]);
        Assert.Equal("2cm", Fmt(section)["marginRight"]);
        Assert.Equal(2, Convert.ToInt32(Fmt(section)["columns"]));
        Assert.Equal("1cm", Fmt(section)["columnSpace"]);
    }
}
