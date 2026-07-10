// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class ExcelConditionalFormattingTests : ExcelTestBase
{
    [Fact]
    public void ConditionalFormattingRuleFamilies_RoundTripThroughGetAndQuery()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            for (var row = 1; row <= 12; row++)
            {
                handler.Add($"/Sheet1/A{row}", "cell", null, new(StringComparer.OrdinalIgnoreCase)
                {
                    ["value"] = row.ToString()
                });
                handler.Add($"/Sheet1/B{row}", "cell", null, new(StringComparer.OrdinalIgnoreCase)
                {
                    ["value"] = row % 2 == 0 ? "error item" : "ok"
                });
            }

            handler.Add("/Sheet1", "databar", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:A12",
                ["min"] = "0",
                ["max"] = "12",
                ["color"] = "4472C4",
                ["negativeColor"] = "FF0000",
                ["axisColor"] = "000000",
                ["axisPosition"] = "middle",
                ["minLength"] = "10",
                ["maxLength"] = "90",
                ["direction"] = "leftToRight",
                ["showValue"] = "false"
            });
            handler.Add("/Sheet1", "colorscale", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:A12",
                ["minColor"] = "F8696B",
                ["midColor"] = "FFEB84",
                ["maxColor"] = "63BE7B",
                ["midpoint"] = "40"
            });
            handler.Add("/Sheet1", "iconset", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:A12",
                ["iconset"] = "3arrows",
                ["reverse"] = "true",
                ["showValue"] = "false"
            });
            handler.Add("/Sheet1", "formulacf", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:A12",
                ["formula"] = "$A1>6",
                ["fill"] = "FFF2CC",
                ["font.bold"] = "true"
            });
            handler.Add("/Sheet1", "cellis", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:A12",
                ["operator"] = "between",
                ["value"] = "3",
                ["value2"] = "9",
                ["fill"] = "D9EAD3"
            });
            handler.Add("/Sheet1", "topn", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:A12",
                ["rank"] = "3",
                ["percent"] = "true",
                ["bottom"] = "true",
                ["fill"] = "D9E2F3"
            });
            handler.Add("/Sheet1", "aboveaverage", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:A12",
                ["aboveAverage"] = "false",
                ["stdDev"] = "1",
                ["equalAverage"] = "true",
                ["fill"] = "EADCF8"
            });
            handler.Add("/Sheet1", "duplicatevalues", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "B1:B12",
                ["fill"] = "F4CCCC"
            });
            handler.Add("/Sheet1", "uniquevalues", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "A1:A12",
                ["fill"] = "D0E0E3"
            });
            handler.Add("/Sheet1", "containstext", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "B1:B12",
                ["text"] = "error",
                ["fill"] = "FCE5CD"
            });
            handler.Add("/Sheet1", "dateoccurring", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "C1:C12",
                ["period"] = "last7Days",
                ["fill"] = "D9D2E9"
            });
            handler.Add("/Sheet1", "cfextended", null, new(StringComparer.OrdinalIgnoreCase)
            {
                ["ref"] = "B1:B12",
                ["type"] = "beginsWith",
                ["text"] = "err",
                ["fill"] = "CFE2F3"
            });

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        AssertCf(readOnly, 1, "dataBar", "A1:A12").Format.Should().Contain(new Dictionary<string, object?>
        {
            ["color"] = "#4472C4",
            ["negativeColor"] = "#FF0000",
            ["axisColor"] = "#000000",
            ["axisPosition"] = "middle",
            ["min"] = "0",
            ["max"] = "12",
            ["minLength"] = 10U,
            ["maxLength"] = 90U,
            ["direction"] = "leftToRight",
            ["showValue"] = false
        });
        AssertCf(readOnly, 2, "colorScale", "A1:A12").Format.Should().Contain(new Dictionary<string, object?>
        {
            ["minColor"] = "#F8696B",
            ["midColor"] = "#FFEB84",
            ["maxColor"] = "#63BE7B",
            ["midpoint"] = "40"
        });
        AssertCf(readOnly, 3, "iconSet", "A1:A12").Format.Should().Contain(new Dictionary<string, object?>
        {
            ["iconset"] = "3Arrows",
            ["reverse"] = true,
            ["showValue"] = false
        });
        AssertCf(readOnly, 4, "formula", "A1:A12").Format.Should().Contain("formula", "$A1>6");
        AssertCf(readOnly, 5, "cellIs", "A1:A12").Format.Should().Contain(new Dictionary<string, object?>
        {
            ["operator"] = "between",
            ["value"] = "3",
            ["value2"] = "9"
        });
        AssertCf(readOnly, 6, "topN", "A1:A12").Format.Should().Contain(new Dictionary<string, object?>
        {
            ["rank"] = 3U,
            ["percent"] = true,
            ["bottom"] = true
        });
        AssertCf(readOnly, 7, "aboveAverage", "A1:A12").Format.Should().Contain(new Dictionary<string, object?>
        {
            ["aboveAverage"] = false,
            ["stdDev"] = 1,
            ["equalAverage"] = true
        });
        AssertCf(readOnly, 8, "duplicateValues", "B1:B12");
        AssertCf(readOnly, 9, "uniqueValues", "A1:A12");
        AssertCf(readOnly, 10, "containsText", "B1:B12").Format.Should().Contain("text", "error");
        AssertCf(readOnly, 11, "timePeriod", "C1:C12").Format.Should().Contain("period", "last7Days");
        AssertCf(readOnly, 12, "beginsWith", "B1:B12").Format.Should().Contain("text", "err");
        readOnly.Validate().Should().BeEmpty();
    }

    private static OfficeCli.Core.DocumentNode AssertCf(OfficeCli.Handlers.ExcelHandler handler, int index, string type, string reference)
    {
        var node = ReadNode(handler, $"/Sheet1/cf[{index}]");
        node.Type.Should().Be("conditionalFormatting");
        node.Format.Should().Contain("type", type);
        node.Format.Should().Contain("ref", reference);
        return node;
    }
}
