using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordFieldParserTests : WordTestBase
{
    [Theory]
    [InlineData("PAGE", true)]
    [InlineData(" page \\h", true)]
    [InlineData("SEQ Figure \\* ROMAN", true)]
    [InlineData("TOC \\o \"1-3\"", true)]
    [InlineData("FORMTEXT", false)]
    [InlineData("FORMCHECKBOX", false)]
    [InlineData("GOTOBUTTON target", false)]
    [InlineData("MACROBUTTON Macro", false)]
    [InlineData("", false)]
    public void IsDynamicFieldInstruction_DistinguishesRenderedAndInteractiveFields(
        string instruction, bool expected)
    {
        Assert.Equal(expected, WordHandler.IsDynamicFieldInstruction(instruction));
    }

    [Fact]
    public void RecalcSeqFields_AppliesNextRepeatResetAndSupportedFormats()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        AddFieldParagraph(handler, "SEQ Figure");
        AddFieldParagraph(handler, @"SEQ Figure \c");
        AddFieldParagraph(handler, @"SEQ Figure \r 5 \* ROMAN");
        AddFieldParagraph(handler, @"SEQ Figure \* alphabetic");

        var patched = handler.RecalcSeqFields();
        var fields = handler.Query("field");

        Assert.Equal(4, patched);
        Assert.Equal(new[] { "1", "1", "V", "f" }, fields.Select(field => field.Text));
    }

    [Fact]
    public void RecalcSeqFields_LeavesDeferredSwitchWithItsExistingCachedText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        AddFieldParagraph(handler, "SEQ Figure");
        AddFieldParagraph(handler, @"SEQ Figure \# ""0.00""", "keep this cache");

        var patched = handler.RecalcSeqFields();
        var fields = handler.Query("field");

        Assert.Equal(1, patched);
        Assert.Equal("1", fields[0].Text);
        Assert.Equal("keep this cache", fields[1].Text);
    }

    private static void AddFieldParagraph(WordHandler handler, string instruction, string? text = null)
    {
        var paragraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var parts = instruction.Split(' ', 2);
        var props = new Dictionary<string, string>
        {
            ["fieldType"] = parts[0].ToLowerInvariant(),
            ["identifier"] = parts.Length > 1 && parts[0].Equals("SEQ", StringComparison.OrdinalIgnoreCase)
                ? parts[1].Split(' ', 2)[0]
                : "Figure"
        };
        if (parts.Length > 1 && parts[0].Equals("SEQ", StringComparison.OrdinalIgnoreCase))
            props["switches"] = parts[1].Contains(' ') ? parts[1][parts[1].IndexOf(' ')..] : "";
        if (text != null) props["text"] = text;
        handler.Add(paragraph, "field", null, props);
    }
}
