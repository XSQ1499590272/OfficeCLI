using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordEquationContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddSetAndRemoveDisplayEquation_ReadsBackFormulaAndMode()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var equationPath = handler.Add("/body", "equation", null, new()
        {
            ["formula"] = "x",
            ["mode"] = "display"
        });

        var equation = handler.Get(equationPath);
        Assert.Equal("/body/oMathPara[1]", equationPath);
        Assert.Equal("equation", equation.Type);
        Assert.Equal("display", Fmt(equation)["mode"]);
        Assert.Equal("x", equation.Text);

        var unsupported = handler.Set(equationPath, new() { ["formula"] = "y" });
        var updated = handler.Get(equationPath);

        Assert.Empty(unsupported);
        Assert.Equal("y", updated.Text);

        handler.Remove(equationPath);

        Assert.Empty(handler.Query("equation"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddInlineEquation_ToParagraph_ReadsBackFormulaAndMode()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Before " });
        var equationPath = handler.Add(paragraphPath, "equation", null, new()
        {
            ["formula"] = "a",
            ["mode"] = "inline"
        });

        var equation = handler.Get(equationPath);

        Assert.EndsWith("/oMath[1]", equationPath);
        Assert.Equal("equation", equation.Type);
        Assert.Equal("inline", Fmt(equation)["mode"]);
        Assert.Equal("a", equation.Text);

        var unsupported = handler.Set(equationPath, new() { ["formula"] = "b" });
        var updated = handler.Get(equationPath);

        Assert.Empty(unsupported);
        Assert.Equal("inline", Fmt(updated)["mode"]);
        Assert.Equal("b", updated.Text);
        Assert.Empty(handler.Validate());
    }
}
