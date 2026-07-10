using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordFormFieldContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddTextFormField_ReadsBackTextDefaultsAndMetadata()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var fieldPath = handler.Add("/body", "formfield", null, new()
        {
            ["name"] = "TextField1",
            ["type"] = "text",
            ["text"] = "Current value",
            ["default"] = "Default value",
            ["maxLength"] = "25",
            ["textType"] = "number",
            ["textFormat"] = "0.00",
            ["enabled"] = "false",
            ["helpText"] = "Enter a number",
            ["statusText"] = "Number field",
            ["calcOnExit"] = "true",
        });

        var field = handler.Get(fieldPath);
        var queried = Assert.Single(handler.Query("formfield"));

        Assert.Equal("formfield", field.Type);
        Assert.Equal("Current value", field.Text);
        Assert.Equal("TextField1", Fmt(field)["name"]);
        Assert.Equal("text", Fmt(field)["type"]);
        Assert.Equal("Default value", Fmt(field)["default"]);
        Assert.Equal(25, Fmt(field)["maxLength"]);
        Assert.Equal("number", Fmt(field)["textType"]);
        Assert.Equal("0.00", Fmt(field)["textFormat"]);
        Assert.Equal(false, Fmt(field)["enabled"]);
        Assert.Equal("Enter a number", Fmt(field)["helpText"]);
        Assert.Equal("Number field", Fmt(field)["statusText"]);
        Assert.Equal(true, Fmt(field)["calcOnExit"]);
        Assert.Equal("TextField1", Fmt(queried)["name"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddCheckboxFormField_ReadsAndUpdatesCurrentState()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var fieldPath = handler.Add("/body", "formfield", null, new()
        {
            ["name"] = "Check1",
            ["type"] = "checkbox",
            ["checked"] = "true",
            ["default"] = "false",
            ["checkBoxSize"] = "24",
        });

        var field = handler.Get(fieldPath);

        Assert.Equal("checkbox", Fmt(field)["type"]);
        Assert.Equal(true, Fmt(field)["checked"]);
        Assert.Equal(false, Fmt(field)["default"]);
        Assert.Equal("24", Fmt(field)["checkBoxSize"]);
        Assert.Equal("true", field.Text);

        handler.Set(fieldPath, new() { ["checked"] = "false" });
        var updated = handler.Get(fieldPath);

        Assert.Equal(false, Fmt(updated)["checked"]);
        Assert.Equal("false", updated.Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddDropdownFormField_ReadsAndUpdatesSelection()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var fieldPath = handler.Add("/body", "formfield", null, new()
        {
            ["name"] = "Choice1",
            ["type"] = "dropdown",
            ["items"] = "Alpha,Beta,Gamma",
            ["result"] = "1",
            ["default"] = "0",
        });

        var field = handler.Get(fieldPath);

        Assert.Equal("dropdown", Fmt(field)["type"]);
        Assert.Equal("Alpha,Beta,Gamma", Fmt(field)["items"]);
        Assert.Equal(1, Fmt(field)["result"]);
        Assert.Equal(0, Fmt(field)["default"]);
        Assert.Equal("Beta", field.Text);

        handler.Set(fieldPath, new() { ["text"] = "Gamma" });
        var updated = handler.Get(fieldPath);

        Assert.Equal(2, Fmt(updated)["result"]);
        Assert.Equal("Gamma", updated.Text);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedAdd_FormFieldWithPathSpecialOrWhitespaceName_DoesNotPersistField()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);

        Assert.Throws<ArgumentException>(() =>
            handler.Add("/body", "formfield", null, new() { ["name"] = "Bad/Name" }));
        Assert.Throws<ArgumentException>(() =>
            handler.Add("/body", "formfield", null, new() { ["name"] = "Bad Name" }));
        Assert.Empty(handler.Query("formfield"));
    }

    [Fact]
    public void SetFormField_InvalidDropdownInputFollowsCurrentCompatibilityBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var fieldPath = handler.Add("/body", "formfield", null, new()
        {
            ["name"] = "ChoiceGuard",
            ["type"] = "dropdown",
            ["items"] = "Alpha,Beta,Gamma",
            ["result"] = "1"
        });

        var textUnsupported = handler.Set(fieldPath, new() { ["text"] = "Missing" });
        Assert.Empty(textUnsupported);
        var afterText = handler.Get(fieldPath);
        Assert.Equal("Missing", afterText.Text);
        Assert.Equal(1, Fmt(afterText)["result"]);

        var resultUnsupported = handler.Set(fieldPath, new() { ["result"] = "99" });
        Assert.Contains("result", resultUnsupported);

        var current = handler.Get(fieldPath);
        Assert.Equal("Missing", current.Text);
        Assert.Equal(1, Fmt(current)["result"]);
        Assert.Empty(handler.Validate());
    }
}
