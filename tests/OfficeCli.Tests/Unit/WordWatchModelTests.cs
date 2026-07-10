using System.Text.Json;
using OfficeCli.Core;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordWatchModelTests
{
    [Fact]
    public void ExtractWordScrollTarget_MapsBodyParagraphTableRowAndCellPaths()
    {
        Assert.Equal("#w-p-2", WatchMessage.ExtractWordScrollTarget("/body/p[2]"));
        Assert.Equal("#w-p-2", WatchMessage.ExtractWordScrollTarget("/body/paragraph[2]"));
        Assert.Equal("#w-table-1", WatchMessage.ExtractWordScrollTarget("/body/table[1]"));
        Assert.Equal("[data-path=\"/body/table[1]/tr[2]\"]",
            WatchMessage.ExtractWordScrollTarget("/body/table[1]/tr[2]"));
        Assert.Equal("[data-path=\"/body/table[1]/tr[2]/tc[3]\"]",
            WatchMessage.ExtractWordScrollTarget("/body/table[1]/tr[2]/tc[3]"));
    }

    [Fact]
    public void ExtractWordScrollTarget_RejectsNonBodyAndInlinePaths()
    {
        Assert.Null(WatchMessage.ExtractWordScrollTarget(null));
        Assert.Null(WatchMessage.ExtractWordScrollTarget(""));
        Assert.Null(WatchMessage.ExtractWordScrollTarget("/footer[1]/p[1]"));
        Assert.Null(WatchMessage.ExtractWordScrollTarget("/body/p[1]/r[1]"));
        Assert.Null(WatchMessage.ExtractWordScrollTarget("/body/table[1]/tr[1]/tc[1]/r[1]"));
    }

    [Fact]
    public void MarkWireModels_PreserveJsonNamesAndVersionedMarkList()
    {
        var response = new MarksResponse
        {
            Version = 7,
            Marks =
            [
                new WatchMark
                {
                    Id = "m1",
                    Path = "/body/p[1]",
                    Find = "r\"TODO\\d+\"",
                    Color = "#FF0000",
                    Note = "review",
                    Tofix = "replace",
                    MatchedText = ["TODO1"],
                    Stale = false,
                    CreatedAt = DateTime.UtcNow
                }
            ]
        };

        var json = JsonSerializer.Serialize(response, WatchMarkJsonContext.Default.MarksResponse);
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.Equal(7, root.GetProperty("version").GetInt32());
        var mark = root.GetProperty("marks")[0];
        Assert.Equal("m1", mark.GetProperty("id").GetString());
        Assert.Equal("/body/p[1]", mark.GetProperty("path").GetString());
        Assert.Equal("TODO1", mark.GetProperty("matched_text")[0].GetString());

        var roundTrip = JsonSerializer.Deserialize(json, WatchMarkJsonContext.Default.MarksResponse);
        Assert.NotNull(roundTrip);
        Assert.Equal(7, roundTrip!.Version);
        Assert.Equal("review", roundTrip.Marks[0].Note);
    }
}
