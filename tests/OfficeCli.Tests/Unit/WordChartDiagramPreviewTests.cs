using System.Reflection;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordChartDiagramPreviewTests : WordTestBase
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    [Fact]
    public void Chart_AddGetQuery_ReadsBackTypeTitleSeriesCategoriesAndSize()
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
    public void Chart_SetSeriesAxisAndRangeBackedData_ReadsBackCurrentSurface()
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
            ["data"] = "Revenue:1,2,3;Cost:4,5,6"
        });

        Assert.Empty(handler.Set($"{chartPath}/series[2]", new()
        {
            ["name"] = "Expense",
            ["values"] = "8,10,13"
        }));
        Assert.Empty(handler.Set($"{chartPath}/axis[@role=value]", new()
        {
            ["title"] = "Amount",
            ["min"] = "0",
            ["max"] = "40",
            ["format"] = "$#,##0"
        }));

        var chart = handler.Get(chartPath);
        var firstSeries = handler.Get($"{chartPath}/series[1]");
        var secondSeries = handler.Get($"{chartPath}/series[2]");
        var axis = handler.Get($"{chartPath}/axis[@role=value]");

        Assert.Equal("1,2,3", Fmt(chart)["categories"]);
        Assert.Equal("Sheet1!$A$1:$A$3", Fmt(chart)["categoriesRef"]);
        Assert.Equal("Revenue:1,2,3", Fmt(chart)["series1"]);
        Assert.Equal("Expense:8,10,13", Fmt(chart)["series2"]);
        Assert.Equal("Revenue", Fmt(firstSeries)["name"]);
        Assert.Equal("1,2,3", Fmt(firstSeries)["values"]);
        Assert.Equal("Sheet1!$B$1:$B$3", Fmt(firstSeries)["valuesRef"]);
        Assert.Equal("Sheet1!$A$1:$A$3", Fmt(firstSeries)["categoriesRef"]);
        Assert.Equal("#FF0000", Fmt(firstSeries)["color"]);
        Assert.Equal("Expense", Fmt(secondSeries)["name"]);
        Assert.Equal("8,10,13", Fmt(secondSeries)["values"]);
        Assert.Equal("value", Fmt(axis)["role"]);
        Assert.Equal("Amount", Fmt(axis)["title"]);
        Assert.Equal("0", Fmt(axis)["min"]);
        Assert.Equal("40", Fmt(axis)["max"]);
        Assert.Equal("$#,##0", Fmt(axis)["format"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void Chart_InvalidSeriesValue_CurrentlyClearsValuesAndLeavesSchemaError()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var chartPath = handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "Jan,Feb,Mar",
            ["data"] = "Revenue:1,2,3"
        });

        var ex = Assert.Throws<CliException>(() =>
            handler.Set($"{chartPath}/series[1]", new() { ["values"] = "1,NaN,3" }));
        var chart = handler.Get(chartPath);
        var validationError = Assert.Single(handler.Validate());

        Assert.Equal("invalid_value", ex.Code);
        Assert.Equal("Revenue:", Fmt(chart)["series1"]);
        Assert.Equal("Schema", validationError.ErrorType);
        Assert.Contains("incomplete content", validationError.Description);
    }

    [Fact]
    public void ChartHelper_ParsesChartTypeSeriesRangesAndColors_CurrentBehavior()
    {
        const BindingFlags Flags = BindingFlags.Static | BindingFlags.NonPublic;
        var chartHelper = typeof(CliException).Assembly.GetType("OfficeCli.Core.ChartHelper", throwOnError: true)!;

        object? Invoke(string name, object arg) =>
            chartHelper.GetMethod(name, Flags)!.Invoke(null, new[] { arg });
        T Field<T>(object value, string name) =>
            (T)value.GetType().GetField(name)!.GetValue(value)!;
        T? Prop<T>(object value, string name) =>
            (T?)value.GetType().GetProperty(name)!.GetValue(value);

        var parsedType = Invoke("ParseChartType", "percentStackedBar3d")!;
        Assert.Equal("bar", Field<string>(parsedType, "Item1"));
        Assert.True(Field<bool>(parsedType, "Item2"));
        Assert.False(Field<bool>(parsedType, "Item3"));
        Assert.True(Field<bool>(parsedType, "Item4"));

        var literalSeries = ((System.Collections.IEnumerable)Invoke("ParseSeriesData", new Dictionary<string, string>
        {
            ["data"] = "Persons (Data year: 2021):1,2,3;Cost:4,5,6"
        })!).Cast<object>().ToList();
        Assert.Equal("Persons (Data year: 2021)", Field<string>(literalSeries[0], "Item1"));
        Assert.Equal(new[] { 1d, 2d, 3d }, Field<double[]>(literalSeries[0], "Item2"));
        Assert.Equal("Cost", Field<string>(literalSeries[1], "Item1"));

        var refSeries = Assert.Single(((System.Collections.IEnumerable)Invoke("ParseSeriesDataExtended", new Dictionary<string, string>
        {
            ["series1.name"] = "Revenue",
            ["series1.values"] = "Sheet1!B1:B3",
            ["series1.categories"] = "Sheet1!A1:A3",
            ["series1.bubbleSize"] = "Sheet1!D1:D3"
        })!).Cast<object>());
        Assert.Equal("Revenue", Prop<string>(refSeries, "Name"));
        Assert.Null(Prop<double[]>(refSeries, "Values"));
        Assert.Equal("Sheet1!$B$1:$B$3", Prop<string>(refSeries, "ValuesRef"));
        Assert.Equal("Sheet1!$A$1:$A$3", Prop<string>(refSeries, "CategoriesRef"));
        Assert.Equal("Sheet1!$D$1:$D$3", Prop<string>(refSeries, "BubbleSizeRef"));

        var categories = Assert.IsType<string[]>(Invoke("ParseCategories", new Dictionary<string, string>
        {
            ["categories"] = "Q1, Q2, Q3"
        }));
        Assert.Equal(new[] { "Q1", "Q2", "Q3" }, categories);
        Assert.Equal("Sheet1!$A$1:$A$3", Invoke("ParseCategoriesRef", new Dictionary<string, string>
        {
            ["categories"] = "Sheet1!A1:A3"
        }));

        var colors = Assert.IsType<string[]>(Invoke("ParseSeriesColors", new Dictionary<string, string>
        {
            ["colors"] = "#111111,#222222,#333333",
            ["series2.color"] = "#ABCDEF"
        }));
        Assert.Equal(new[] { "#111111", "#ABCDEF", "#333333" }, colors);
    }

    [Fact]
    public void Diagram_NativeAddSetRemoveAndMissingSource_CurrentBehavior()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var groupPath = handler.Add("/body", "diagram", null, new()
        {
            ["render"] = "native",
            ["mermaid"] = "flowchart TD; A[Start] --> B[Done]",
            ["width"] = "8cm"
        });

        var group = handler.Get(groupPath);
        var firstNode = handler.Get("/body/textbox[1]/p[1]");

        Assert.Equal("/body/group[1]", groupPath);
        Assert.Equal("group", group.Type);
        Assert.True(Fmt(group).ContainsKey("x"));
        Assert.True(Fmt(group).ContainsKey("y"));
        Assert.True(Fmt(group).ContainsKey("width"));
        Assert.True(Fmt(group).ContainsKey("height"));
        Assert.Equal("Start", firstNode.Text);
        Assert.Throws<ArgumentException>(() => handler.Get("/body/textbox[2]/p[1]"));

        Assert.Empty(handler.Set(groupPath, new()
        {
            ["width"] = "6cm",
            ["height"] = "3cm"
        }));
        var resized = handler.Get(groupPath);
        Assert.Equal("6cm", Fmt(resized)["width"]);
        Assert.Equal("3cm", Fmt(resized)["height"]);

        handler.Remove(groupPath);
        Assert.Throws<ArgumentException>(() => handler.Get(groupPath));
        Assert.Empty(handler.Validate());

        var ex = Assert.Throws<ArgumentException>(() =>
            handler.Add("/body", "diagram", null, new() { ["render"] = "native" }));
        Assert.Contains("diagram requires", ex.Message);
        Assert.Throws<ArgumentException>(() => handler.Get("/body/group[1]"));
    }

    [Fact]
    public void HtmlPreview_RendersParagraphTablePictureDataUriAndAltText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Preview paragraph" });
        handler.Add("/body", "table", null, new() { ["data"] = "A,B" });
        handler.Add("/body", "picture", null, new()
        {
            ["src"] = TinyPngDataUri,
            ["name"] = "preview.png",
            ["alt"] = "Preview image"
        });

        var html = handler.ViewAsHtml();

        Assert.Contains("<!DOCTYPE html>", html);
        Assert.Contains("Preview paragraph", html);
        Assert.Contains("<table", html);
        Assert.Contains(">A<", html);
        Assert.Contains(">B<", html);
        Assert.Contains("<img ", html);
        Assert.Contains("src=\"data:image/png;base64,", html);
        Assert.Contains("alt=\"Preview image\"", html);
    }
}
