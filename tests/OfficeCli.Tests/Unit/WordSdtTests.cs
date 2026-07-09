using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordSdtTests : WordTestBase
{
    [Fact]
    public void AddDropdownComboAndDateSdt_ReadBackTypeSpecificProperties()
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
        var datePath = handler.Add("/body", "sdt", null, new()
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

        var dropdown = handler.Get(dropdownPath);
        var combo = handler.Get(comboPath);
        var date = handler.Get(datePath);

        Assert.Equal("dropdown", Fmt(dropdown)["type"]);
        Assert.Equal("Status", Fmt(dropdown)["alias"]);
        Assert.Equal("status", Fmt(dropdown)["tag"]);
        Assert.Equal("Draft|DRAFT,Final|FINAL", Fmt(dropdown)["items"]);
        Assert.Equal("FINAL", Fmt(dropdown)["dropDown.lastValue"]);
        Assert.Equal("combobox", Fmt(combo)["type"]);
        Assert.Equal("Small|S,Large|L", Fmt(combo)["items"]);
        Assert.Equal("L", Fmt(combo)["comboBox.lastValue"]);
        Assert.Equal("date", Fmt(date)["type"]);
        Assert.Equal("yyyy-MM-dd", Fmt(date)["format"]);
        Assert.Equal("2026-01-02T00:00:00Z", Fmt(date)["date.fullDate"]);
        Assert.Equal("gregorian", Fmt(date)["date.calendar"]);
        Assert.Equal("en-US", Fmt(date)["date.lid"]);
        Assert.Equal("dateTime", Fmt(date)["date.storeMappedDataAs"]);
        Assert.Equal(true, Fmt(date)["placeholder"]);
        Assert.Equal("DefaultPlaceholder", Fmt(date)["placeholderText"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddGroupPictureRichTextAndStableSdtPath_ReadBackImplementedMarkers()
    {
        var path = CreateBlankDocx();
        string textPath;
        object? sdtId;

        using (var handler = new WordHandler(path, editable: true))
        {
            var groupPath = handler.Add("/body", "sdt", null, new() { ["type"] = "group", ["text"] = "Group" });
            var picturePath = handler.Add("/body", "sdt", null, new() { ["type"] = "picture", ["text"] = "Picture" });
            var richPath = handler.Add("/body", "sdt", null, new() { ["type"] = "richtext", ["text"] = "Rich" });
            textPath = handler.Add("/body", "sdt", null, new() { ["type"] = "text", ["text"] = "stable sdt" });
            sdtId = Fmt(handler.Get(textPath))["id"];

            Assert.Equal("group", Fmt(handler.Get(groupPath))["type"]);
            Assert.Equal("picture", Fmt(handler.Get(picturePath))["type"]);
            Assert.Equal("richtext", Fmt(handler.Get(richPath))["type"]);
            Assert.Empty(handler.Validate());
        }

        using var reopened = new WordHandler(path, editable: false);
        var stable = reopened.Get($"/body/sdt[@sdtId={sdtId}]");

        Assert.Equal("sdt", stable.Type);
        Assert.Equal("stable sdt", stable.Text);
    }
}
