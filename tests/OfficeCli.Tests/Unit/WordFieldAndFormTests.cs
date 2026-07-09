using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordFieldAndFormTests : WordTestBase
{
    [Fact]
    public void AddField_ReadsBackInstructionResultAndFieldChars()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(paragraphPath, "field", null, new()
        {
            ["fieldType"] = "page",
            ["text"] = "7",
            ["fldLock"] = "true"
        });

        var field = Assert.Single(handler.Query("field"));
        var fieldChars = handler.Query("fieldChar");

        Assert.Equal("7", field.Text);
        Assert.Equal("PAGE", Fmt(field)["instruction"]);
        Assert.Equal("page", Fmt(field)["fieldType"]);
        Assert.Equal(true, Fmt(field)["evaluated"]);
        Assert.Equal(true, Fmt(field)["fldLock"]);
        Assert.Contains(fieldChars, node => Equals(Fmt(node)["fieldCharType"], "begin"));
        Assert.Contains(fieldChars, node => Equals(Fmt(node)["fieldCharType"], "separate"));
        Assert.Contains(fieldChars, node => Equals(Fmt(node)["fieldCharType"], "end"));
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FieldVirtualPath_RemoveIsRejectedCurrentBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(paragraphPath, "field", null, new()
        {
            ["fieldType"] = "page",
            ["text"] = "7"
        });

        var field = Assert.Single(handler.Query("field"));
        var error = Assert.Throws<ArgumentException>(() => handler.Remove(field.Path));

        Assert.Equal("/field[1]", field.Path);
        Assert.Contains("Path not found: /field[1]", error.Message);
    }

    [Fact]
    public void RemoveInstrTextPath_DropsInstructionAndCollapsedField()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(paragraphPath, "field", null, new()
        {
            ["fieldType"] = "page",
            ["text"] = "7"
        });

        var instrText = Assert.Single(handler.Query("instrText"));
        handler.Remove(instrText.Path);

        Assert.Empty(handler.Query("instrText"));
        Assert.Empty(handler.Query("field"));
        Assert.Equal(3, handler.Query("fieldChar").Count);
    }

    [Fact]
    public void RemoveFieldCharPath_DropsOnlyTargetMarkerAndCollapsedField()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(paragraphPath, "field", null, new()
        {
            ["fieldType"] = "page",
            ["text"] = "7"
        });

        var begin = handler.Query("fieldChar")
            .Single(node => Equals(Fmt(node)["fieldCharType"], "begin"));
        handler.Remove(begin.Path);

        var remainingTypes = handler.Query("fieldChar")
            .Select(node => Fmt(node)["fieldCharType"])
            .ToList();
        Assert.DoesNotContain("begin", remainingTypes);
        Assert.Contains("separate", remainingTypes);
        Assert.Contains("end", remainingTypes);
        Assert.Empty(handler.Query("field"));
    }

    [Fact]
    public void SeqFields_CacheCurrentBodyOrderAndResetRomanBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var firstParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var secondParagraph = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });

        handler.Add(firstParagraph, "field", null, new()
        {
            ["fieldType"] = "seq",
            ["identifier"] = "Figure",
            ["switches"] = @"\r 5 \* ROMAN"
        });
        handler.Add(secondParagraph, "field", null, new()
        {
            ["fieldType"] = "seq",
            ["identifier"] = "Figure",
            ["switches"] = @"\* ROMAN"
        });

        var fields = handler.Query("field");

        Assert.Equal("V", fields[0].Text);
        Assert.Equal("II", fields[1].Text);
        Assert.Equal(@"SEQ Figure \r 5 \* ROMAN", Fmt(fields[0])["instruction"]?.ToString()?.Trim());
        Assert.Equal(@"SEQ Figure \* ROMAN", Fmt(fields[1])["instruction"]?.ToString()?.Trim());
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FormFields_ReadBackTextCheckboxDropdownAndRejectBadNames()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var textPath = handler.Add("/body", "formfield", null, new()
        {
            ["name"] = "Text1",
            ["type"] = "text",
            ["text"] = "Current",
            ["default"] = "Default",
            ["maxLength"] = "25"
        });
        var checkboxPath = handler.Add("/body", "formfield", null, new()
        {
            ["name"] = "Check1",
            ["type"] = "checkbox",
            ["checked"] = "true",
            ["default"] = "false"
        });
        var dropdownPath = handler.Add("/body", "formfield", null, new()
        {
            ["name"] = "Choice1",
            ["type"] = "dropdown",
            ["items"] = "Alpha,Beta,Gamma",
            ["result"] = "1"
        });

        var text = handler.Get(textPath);
        var checkbox = handler.Get(checkboxPath);
        var dropdown = handler.Get(dropdownPath);

        Assert.Equal("Current", text.Text);
        Assert.Equal("Default", Fmt(text)["default"]);
        Assert.Equal(25, Fmt(text)["maxLength"]);
        Assert.Equal(true, Fmt(checkbox)["checked"]);
        Assert.Equal(false, Fmt(checkbox)["default"]);
        Assert.Equal("Beta", dropdown.Text);
        Assert.Equal("Alpha,Beta,Gamma", Fmt(dropdown)["items"]);
        Assert.Throws<ArgumentException>(() =>
            handler.Add("/body", "formfield", null, new() { ["name"] = "Bad/Name" }));
        Assert.Empty(handler.Validate());
    }
}
