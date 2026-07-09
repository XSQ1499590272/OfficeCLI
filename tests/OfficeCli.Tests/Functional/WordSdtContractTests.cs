using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordSdtContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddDropdownAndComboBox_ReadsBackItemsAndCurrentSelection()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var dropdownPath = handler.Add("/body", "sdt", null, new()
        {
            ["type"] = "dropdown",
            ["alias"] = "Status",
            ["tag"] = "status",
            ["items"] = "Draft|DRAFT,Final|FINAL",
            ["dropDown.lastValue"] = "FINAL",
            ["text"] = "Final"
        });
        var comboPath = handler.Add("/body", "sdt", null, new()
        {
            ["type"] = "combobox",
            ["items"] = "Small|S,Large|L",
            ["comboBox.lastValue"] = "L",
            ["text"] = "Large"
        });

        var dropdown = handler.Get(dropdownPath);
        var combo = handler.Get(comboPath);

        Assert.Equal("dropdown", Fmt(dropdown)["type"]);
        Assert.Equal("Status", Fmt(dropdown)["alias"]);
        Assert.Equal("status", Fmt(dropdown)["tag"]);
        Assert.Equal("Draft|DRAFT,Final|FINAL", Fmt(dropdown)["items"]);
        Assert.Equal("FINAL", Fmt(dropdown)["dropDown.lastValue"]);
        Assert.Equal("Final", dropdown.Text);
        Assert.Equal("combobox", Fmt(combo)["type"]);
        Assert.Equal("Small|S,Large|L", Fmt(combo)["items"]);
        Assert.Equal("L", Fmt(combo)["comboBox.lastValue"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddDateSdt_ReadsBackDateSpecificProperties()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var sdtPath = handler.Add("/body", "sdt", null, new()
        {
            ["type"] = "date",
            ["format"] = "yyyy-MM-dd",
            ["date.fullDate"] = "2026-01-02T00:00:00Z",
            ["date.calendar"] = "gregorian",
            ["date.lid"] = "en-US",
            ["date.storeMappedDataAs"] = "dateTime",
            ["placeholder"] = "true",
            ["placeholderText"] = "DefaultPlaceholder",
            ["text"] = "2026-01-02"
        });

        var sdt = handler.Get(sdtPath);

        Assert.Equal("date", Fmt(sdt)["type"]);
        Assert.Equal("yyyy-MM-dd", Fmt(sdt)["format"]);
        Assert.Equal("2026-01-02T00:00:00Z", Fmt(sdt)["date.fullDate"]);
        Assert.Equal("gregorian", Fmt(sdt)["date.calendar"]);
        Assert.Equal("en-US", Fmt(sdt)["date.lid"]);
        Assert.Equal("dateTime", Fmt(sdt)["date.storeMappedDataAs"]);
        Assert.Equal(true, Fmt(sdt)["placeholder"]);
        Assert.Equal("DefaultPlaceholder", Fmt(sdt)["placeholderText"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddGroupPictureAndRichTextSdt_ReadsBackImplementedTypeMarkers()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var groupPath = handler.Add("/body", "sdt", null, new() { ["type"] = "group", ["text"] = "Group" });
        var picturePath = handler.Add("/body", "sdt", null, new() { ["type"] = "picture", ["text"] = "Picture" });
        var richPath = handler.Add("/body", "sdt", null, new() { ["type"] = "richtext", ["text"] = "Rich" });

        Assert.Equal("group", Fmt(handler.Get(groupPath))["type"]);
        Assert.Equal("picture", Fmt(handler.Get(picturePath))["type"]);
        Assert.Equal("richtext", Fmt(handler.Get(richPath))["type"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void StableSdtIdPath_NavigatesAfterReopen()
    {
        var path = CreateBlankDocx();
        string sdtPath;
        object? sdtId;

        using (var handler = new WordHandler(path, editable: true))
        {
            sdtPath = handler.Add("/body", "sdt", null, new()
            {
                ["type"] = "text",
                ["text"] = "stable sdt"
            });
            sdtId = Fmt(handler.Get(sdtPath))["id"];
        }

        using var reopened = new WordHandler(path, editable: false);
        var stablePath = $"/body/sdt[@sdtId={sdtId}]";
        var sdt = reopened.Get(stablePath);

        Assert.Equal("sdt", sdt.Type);
        Assert.Equal("stable sdt", sdt.Text);
    }
}
