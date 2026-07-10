// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using FluentAssertions;
using OfficeCli.Core;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class ExcelChartUnitTests
{
    [Fact]
    public void ChartHelper_ParseHelpers_CoverAliasesAndReferenceNormalization()
    {
        var (kind, is3D, stacked, percentStacked) = ChartHelper.ParseChartType("percentStackedColumn3D");

        kind.Should().Be("column");
        is3D.Should().BeTrue();
        stacked.Should().BeFalse();
        percentStacked.Should().BeTrue();
        ChartHelper.IsRangeReference("A1:B3").Should().BeTrue();
        ChartHelper.NormalizeRangeReference("A1:B3", "Sheet1").Should().Be("Sheet1!$A$1:$B$3");
        ChartHelper.NormalizeCellReference("'Quarterly Data'!B2").Should().Be("'Quarterly Data'!$B$2");
        ChartHelper.ParseLegendPosition("TOP_RIGHT").Should().Be(C.LegendPositionValues.TopRight);

        ChartHelper.TryParseSeriesDottedKey("series2.point3.color", out var seriesIndex, out var property)
            .Should().BeTrue();
        seriesIndex.Should().Be(2);
        property.Should().Be("point3.color");
    }

    [Fact]
    public void TrendlineHelpers_BuildDefaultsAndKeepSchemaOrderWhenApplyingOptions()
    {
        var movingAverage = ChartHelper.BuildTrendline("movingAvg");
        movingAverage.GetFirstChild<C.TrendlineType>()!.Val!.Value.Should().Be(C.TrendlineValues.MovingAverage);
        movingAverage.GetFirstChild<C.Period>()!.Val!.Value.Should().Be(2u);

        var polynomial = ChartHelper.BuildTrendline("poly");
        polynomial.GetFirstChild<C.PolynomialOrder>()!.Val!.Value.Should().Be((byte)2);

        ChartHelper.ApplyTrendlineOptions(polynomial, "name", "Forecast");
        ChartHelper.ApplyTrendlineOptions(polynomial, "dispeq", "true");
        ChartHelper.ApplyTrendlineOptions(polynomial, "disprsqr", "true");
        ChartHelper.ApplyTrendlineOptions(polynomial, "forward", "1.5");

        polynomial.GetFirstChild<C.TrendlineName>()!.Text.Should().Be("Forecast");
        polynomial.GetFirstChild<C.DisplayEquation>()!.Val!.Value.Should().BeTrue();
        polynomial.GetFirstChild<C.DisplayRSquaredValue>()!.Val!.Value.Should().BeTrue();
        polynomial.GetFirstChild<C.Forward>()!.Val!.Value.Should().Be(1.5d);

        polynomial.ChildElements.Select(e => e.LocalName).Should().ContainInOrder(
            "name",
            "trendlineType",
            "order",
            "dispRSqr",
            "dispEq",
            "trendlineLbl",
            "forward");
    }

    [Fact]
    public void ErrorBarHelpers_SupportNumericShorthandDirectionalAndCustomLiterals()
    {
        var shorthand = ChartHelper.BuildErrorBars("5");
        shorthand.GetFirstChild<C.ErrorBarType>()!.Val!.Value.Should().Be(C.ErrorBarValues.Both);
        shorthand.GetFirstChild<C.ErrorBarValueType>()!.Val!.Value.Should().Be(C.ErrorValues.FixedValue);
        shorthand.Elements<C.Plus>().Single().Descendants<C.NumericValue>().Single().Text.Should().Be("5");
        shorthand.Elements<C.Minus>().Single().Descendants<C.NumericValue>().Single().Text.Should().Be("5");

        var directional = ChartHelper.BuildErrorBars("minus:percent:10");
        directional.GetFirstChild<C.ErrorBarType>()!.Val!.Value.Should().Be(C.ErrorBarValues.Minus);
        directional.GetFirstChild<C.ErrorBarValueType>()!.Val!.Value.Should().Be(C.ErrorValues.Percentage);
        directional.Elements<C.Plus>().Single().Descendants<C.NumericValue>().Single().Text.Should().Be("10");

        var custom = ChartHelper.BuildErrorBars("cust:plus:1,2,3:");
        custom.GetFirstChild<C.ErrorBarType>()!.Val!.Value.Should().Be(C.ErrorBarValues.Plus);
        custom.GetFirstChild<C.ErrorBarValueType>()!.Val!.Value.Should().Be(C.ErrorValues.Custom);
        custom.Elements<C.Plus>().Single().Descendants<C.NumericPoint>()
            .Select(point => point.GetFirstChild<C.NumericValue>()!.Text)
            .Should().Equal("1", "2", "3");
        custom.Elements<C.Minus>().Should().BeEmpty();
    }

    [Fact]
    public void HandleSeriesDottedProperty_PreservesMarkerStateAndDeduplicatesTrendlines()
    {
        var lineChart = new C.LineChart(new C.Grouping { Val = C.GroupingValues.Standard });
        var series = new C.LineChartSeries(new C.Index { Val = 0u }, new C.Order { Val = 0u });
        lineChart.Append(series);

        ChartHelper.HandleSeriesDottedProperty(series, "marker", "circle:9").Should().BeTrue();
        ChartHelper.HandleSeriesDottedProperty(series, "marker.color", "FF0000").Should().BeTrue();
        ChartHelper.HandleSeriesDottedProperty(series, "marker.size", "7").Should().BeTrue();
        ChartHelper.HandleSeriesDottedProperty(series, "trendline", "linear;poly:4").Should().BeTrue();
        ChartHelper.HandleSeriesDottedProperty(series, "trendline", "linear").Should().BeTrue();

        var marker = series.GetFirstChild<C.Marker>()!;
        marker.GetFirstChild<C.Symbol>()!.Val!.Value.Should().Be(C.MarkerStyleValues.Circle);
        marker.GetFirstChild<C.Size>()!.Val!.Value.Should().Be((byte)7);
        marker.GetFirstChild<C.ChartShapeProperties>()!
            .GetFirstChild<A.SolidFill>()!
            .GetFirstChild<A.RgbColorModelHex>()!
            .Val!.Value.Should().Be("FF0000");

        var trendlines = series.Elements<C.Trendline>().ToList();
        trendlines.Should().HaveCount(2);
        trendlines.Select(tl => tl.GetFirstChild<C.TrendlineType>()!.Val!.Value)
            .Should().Contain([C.TrendlineValues.Linear, C.TrendlineValues.Polynomial]);
        trendlines.Single(tl => tl.GetFirstChild<C.TrendlineType>()!.Val!.Value == C.TrendlineValues.Polynomial)
            .GetFirstChild<C.PolynomialOrder>()!.Val!.Value.Should().Be((byte)4);
    }

    [Fact]
    public void BuildExtendedChartSpace_WiresEmbeddedWorkbookReferencesAndPerPointColors()
    {
        var chartSpace = ChartExBuilder.BuildExtendedChartSpace(
            "funnel",
            "Pipeline",
            ["Q1", "Q2", "Q3"],
            [("Revenue", new[] { 10d, 20d, 30d })],
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["colors"] = "FF0000,00FF00,0000FF",
                ["dataLabels"] = "true",
                ["datalabels.numfmt"] = "#,##0"
            });

        var chartData = chartSpace.GetFirstChild<CX.ChartData>()!;
        chartData.ChildElements.First().Should().BeOfType<CX.ExternalData>();
        ((CX.ExternalData)chartData.ChildElements.First()).Id!.Value.Should().Be("rId1");

        var formulas = chartData.Descendants<CX.Formula>().Select(formula => formula.Text).ToList();
        formulas.Should().Contain("Sheet1!$A$2:$A$4");
        formulas.Should().Contain("Sheet1!$B$2:$B$4");

        var series = chartSpace.Descendants<CX.Series>().Should().ContainSingle().Subject;
        series.GetFirstChild<CX.Text>()!.Descendants<CX.Formula>().Single().Text.Should().Be("Sheet1!$B$1");
        series.Elements<CX.DataPoint>().Should().HaveCount(3);
        series.GetFirstChild<CX.DataLabels>()!.NumberFormat!.FormatCode!.Value.Should().Be("#,##0");
    }

    [Fact]
    public void ExtractChartInfo_ReadsPointColorsDeletedLabelsMarkersTrendlinesErrorBarsAndReferenceLines()
    {
        var plotArea = new C.PlotArea();
        var lineChart = new C.LineChart(
            new C.Grouping { Val = C.GroupingValues.Standard },
            new C.VaryColors { Val = false });
        plotArea.Append(lineChart);

        var series = new C.LineChartSeries(
            new C.Index { Val = 0u },
            new C.Order { Val = 0u },
            new C.SeriesText(new C.NumericValue("Revenue")),
            new C.Marker(
                new C.Symbol { Val = C.MarkerStyleValues.Diamond },
                new C.Size { Val = 9 },
                new C.ChartShapeProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = "FFD966" }),
                    new A.Outline(
                        new A.SolidFill(new A.RgbColorModelHex { Val = "C55A11" })))),
            new C.DataLabels(
                new C.DataLabel(
                    new C.Index { Val = 1u },
                    new C.Delete { Val = true })),
            new C.DataPoint(
                new C.Index { Val = 1u },
                new C.ChartShapeProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }))),
            new C.Trendline(
                new C.TrendlineType { Val = C.TrendlineValues.Polynomial },
                new C.PolynomialOrder { Val = 3 },
                new C.Forward { Val = 1.5d },
                new C.Backward { Val = 0.5d },
                new C.Intercept { Val = 2d },
                new C.DisplayEquation { Val = true },
                new C.DisplayRSquaredValue { Val = true },
                new C.ChartShapeProperties(
                    new A.Outline(
                        new A.SolidFill(new A.RgbColorModelHex { Val = "4472C4" }),
                        new A.PresetDash { Val = A.PresetLineDashValues.SystemDash }))),
            new C.ErrorBars(
                new C.ErrorDirection { Val = C.ErrorBarDirectionValues.Y },
                new C.ErrorBarType { Val = C.ErrorBarValues.Both },
                new C.ErrorBarValueType { Val = C.ErrorValues.FixedValue },
                new C.Plus(
                    new C.NumberLiteral(
                        new C.FormatCode("General"),
                        new C.PointCount { Val = 1u },
                        new C.NumericPoint(new C.NumericValue("5")) { Index = 0u })),
                new C.Minus(
                    new C.NumberLiteral(
                        new C.FormatCode("General"),
                        new C.PointCount { Val = 1u },
                        new C.NumericPoint(new C.NumericValue("5")) { Index = 0u })),
                new C.ChartShapeProperties(
                    new A.Outline(
                        new A.SolidFill(new A.RgbColorModelHex { Val = "00B050" }))
                    {
                        Width = (int)Math.Round(1.5 * EmuConverter.EmuPerPoint)
                    })),
            new C.CategoryAxisData(
                new C.StringLiteral(
                    new C.PointCount { Val = 3u },
                    new C.StringPoint(new C.NumericValue("Jan")) { Index = 0u },
                    new C.StringPoint(new C.NumericValue("Feb")) { Index = 1u },
                    new C.StringPoint(new C.NumericValue("Mar")) { Index = 2u })),
            new C.Values(
                new C.NumberLiteral(
                    new C.FormatCode("General"),
                    new C.PointCount { Val = 3u },
                    new C.NumericPoint(new C.NumericValue("10")) { Index = 0u },
                    new C.NumericPoint(new C.NumericValue("20")) { Index = 1u },
                    new C.NumericPoint(new C.NumericValue("30")) { Index = 2u })));
        lineChart.Append(series);

        var chart = new C.Chart(plotArea);
        ChartHelper.AddReferenceLine(chart, "50:00AA00:2:dash:Target", removeExisting: false);

        var info = ChartSvgRenderer.ExtractChartInfo(plotArea, chart);

        info.Series.Should().ContainSingle();
        info.Series.Single().name.Should().Be("Revenue");
        info.PerPointColors.Should().HaveCount(2);
        info.PerPointColors[0][1].Should().Be("#FF0000");
        info.PerPointColors[1].Should().BeEmpty();
        info.PerPointDeletedLabels.Should().HaveCount(2);
        info.PerPointDeletedLabels[0].Should().Contain(1);
        info.PerPointDeletedLabels[1].Should().BeEmpty();
        info.MarkerShapes.Should().Equal("diamond", "none");
        info.MarkerSizes.Should().Equal(9, 5);
        info.MarkerFillColors.Should().Equal(new string?[] { "#FFD966", null });
        info.MarkerLineColors.Should().Equal(new string?[] { "#C55A11", null });

        info.Trendlines.Should().HaveCount(2);
        info.Trendlines[0]!.Type.Should().Be("poly");
        info.Trendlines[0]!.Order.Should().Be(3);
        info.Trendlines[0]!.Forward.Should().Be(1.5d);
        info.Trendlines[0]!.Backward.Should().Be(0.5d);
        info.Trendlines[0]!.Intercept.Should().Be(2d);
        info.Trendlines[0]!.DisplayEquation.Should().BeTrue();
        info.Trendlines[0]!.DisplayRSquared.Should().BeTrue();
        info.Trendlines[0]!.Color.Should().Be("4472C4");
        info.Trendlines[0]!.Dash.Should().Be("sysDash");
        info.Trendlines[1].Should().BeNull();

        info.ErrorBars.Should().HaveCount(2);
        info.ErrorBars[0]!.BarType.Should().Be("both");
        info.ErrorBars[0]!.Direction.Should().Be("y");
        info.ErrorBars[0]!.ValueType.Should().Be("fixedVal");
        info.ErrorBars[0]!.Value.Should().Be(5d);
        info.ErrorBars[0]!.Color.Should().Be("00B050");
        info.ErrorBars[0]!.Width.Should().Be(1.5d);
        info.ErrorBars[1].Should().BeNull();

        info.ReferenceLines.Should().ContainSingle();
        info.ReferenceLines[0].Name.Should().Be("Target");
        info.ReferenceLines[0].Value.Should().Be(50d);
        info.ReferenceLines[0].Color.Should().Be("00AA00");
        Math.Round(info.ReferenceLines[0].WidthPt, 1).Should().Be(2d);
        info.ReferenceLines[0].Dash.Should().Be("sysDash");
    }
}
