using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordChartContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddChart_ReadsBackTypeTitleCategoriesSeriesAndSize()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var chartPath = handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Quarterly result",
            ["categories"] = "Q1,Q2,Q3",
            ["data"] = "Revenue:10,20,30;Cost:7,9,12",
            ["width"] = "8cm",
            ["height"] = "5cm"
        });

        var chart = handler.Get(chartPath, depth: 1);
        var queried = Assert.Single(handler.Query("chart"));
        var firstSeries = handler.Get($"{chartPath}/series[1]");

        Assert.Equal("/chart[1]", chartPath);
        Assert.Equal("chart", chart.Type);
        Assert.Equal(chartPath, queried.Path);
        Assert.Equal("column", Fmt(chart)["chartType"]);
        Assert.Equal("Quarterly result", Fmt(chart)["title"]);
        Assert.Equal("Q1,Q2,Q3", Fmt(chart)["categories"]);
        Assert.Equal(2, Fmt(chart)["seriesCount"]);
        Assert.Equal("Revenue:10,20,30", Fmt(chart)["series1"]);
        Assert.Equal("Cost:7,9,12", Fmt(chart)["series2"]);
        Assert.Equal("8.0cm", Fmt(chart)["width"]);
        Assert.Equal("5.0cm", Fmt(chart)["height"]);
        Assert.Equal(2, chart.Children.Count);
        Assert.Equal("series", firstSeries.Type);
        Assert.Equal("Revenue", Fmt(firstSeries)["name"]);
        Assert.Equal("10,20,30", Fmt(firstSeries)["values"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetChartSeriesNameAndValues_ReadsBackUpdatedSeriesOnly()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var chartPath = handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "line",
            ["categories"] = "Jan,Feb,Mar",
            ["data"] = "Revenue:1,2,3;Cost:4,5,6"
        });

        var unsupported = handler.Set($"{chartPath}/series[2]", new()
        {
            ["name"] = "Expense",
            ["values"] = "8,10,13"
        });

        var chart = handler.Get(chartPath, depth: 1);
        var firstSeries = handler.Get($"{chartPath}/series[1]");
        var secondSeries = handler.Get($"{chartPath}/series[2]");

        Assert.Empty(unsupported);
        Assert.Equal("Revenue:1,2,3", Fmt(chart)["series1"]);
        Assert.Equal("Expense:8,10,13", Fmt(chart)["series2"]);
        Assert.Equal("Revenue", Fmt(firstSeries)["name"]);
        Assert.Equal("1,2,3", Fmt(firstSeries)["values"]);
        Assert.Equal("Expense", Fmt(secondSeries)["name"]);
        Assert.Equal("8,10,13", Fmt(secondSeries)["values"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void FailedSetChartSeriesValuesWithInvalidNumber_CurrentlyClearsSeriesValues()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var chartPath = handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "Jan,Feb,Mar",
            ["data"] = "Revenue:1,2,3"
        });

        var ex = Assert.Throws<OfficeCli.Core.CliException>(() =>
            handler.Set($"{chartPath}/series[1]", new() { ["values"] = "1,NaN,3" }));
        var chart = handler.Get(chartPath);
        var validationError = Assert.Single(handler.Validate());

        Assert.Equal("invalid_value", ex.Code);
        Assert.Equal("Revenue:", Fmt(chart)["series1"]);
        Assert.Equal("Schema", validationError.ErrorType);
        Assert.Contains("incomplete content", validationError.Description);
    }

    [Fact]
    public void SetChartValueAxis_ReadsBackTitleScaleAndNumberFormat()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var chartPath = handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "Jan,Feb,Mar",
            ["data"] = "Revenue:10,20,30"
        });

        var axisPath = $"{chartPath}/axis[@role=value]";
        var unsupported = handler.Set(axisPath, new()
        {
            ["title"] = "Amount",
            ["min"] = "0",
            ["max"] = "40",
            ["format"] = "$#,##0"
        });

        var axis = handler.Get(axisPath);
        var fmt = Fmt(axis);

        Assert.Empty(unsupported);
        Assert.Equal("axis", axis.Type);
        Assert.Equal("value", fmt["role"]);
        Assert.Equal("Amount", fmt["title"]);
        Assert.Equal("0", fmt["min"]);
        Assert.Equal("40", fmt["max"]);
        Assert.Equal("$#,##0", fmt["format"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void AddChart_WithRangeBackedSeries_ReadsBackRefsCachedValuesAndColor()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var chartPath = handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "Sheet1!A1:A3",
            ["series1.name"] = "Revenue",
            ["series1.values"] = "Sheet1!B1:B3",
            ["series1.color"] = "#FF0000",
            ["data"] = "Revenue:1,2,3"
        });

        var chart = handler.Get(chartPath);
        var series = handler.Get($"{chartPath}/series[1]");

        Assert.Equal("Revenue:1,2,3", Fmt(chart)["series1"]);
        Assert.Equal("1,2,3", Fmt(chart)["categories"]);
        Assert.Equal("Sheet1!$A$1:$A$3", Fmt(chart)["categoriesRef"]);
        Assert.Equal("Revenue", Fmt(series)["name"]);
        Assert.Equal("1,2,3", Fmt(series)["values"]);
        Assert.Equal("Sheet1!$B$1:$B$3", Fmt(series)["valuesRef"]);
        Assert.Equal("Sheet1!$A$1:$A$3", Fmt(series)["categoriesRef"]);
        Assert.Equal("#FF0000", Fmt(series)["color"]);
        Assert.Empty(handler.Validate());
    }
}
