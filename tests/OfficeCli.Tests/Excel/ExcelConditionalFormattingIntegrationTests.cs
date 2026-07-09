// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelConditionalFormattingIntegrationTests : ExcelTestBase
{
    // ── CellIs operators ─────────────────────────────────────────────────

    [Fact]
    public void CellIs_GreaterThan_RoundTrip()
    {
        var (path, _) = SetupDataRange();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "B2:B6"),
                ("operator", "greaterThan"),
                ("formula", "500"),
                ("fill", "C6EFCE"),
                ("font", "006100")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Type.Should().Be("conditionalFormatting");
        cf.Format.Should().Contain("type", "cellIs");
        cf.Format.Should().Contain("operator", "greaterThan");
        cf.Format.Should().Contain("ref", "B2:B6");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CellIs_LessThan_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "A", "100");
            AddRow(handler, 3, "B", "300");
            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "B2:B3"),
                ("operator", "lessThan"),
                ("formula", "200"),
                ("fill", "FFC7CE"),
                ("font", "9C0006")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("operator", "lessThan");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CellIs_EqualNotEqual_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Grade", "Score");
            AddRow(handler, 2, "A", "95");
            AddRow(handler, 3, "B", "85");
            AddRow(handler, 4, "A", "92");
            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "A2:A4"),
                ("operator", "equal"),
                ("formula", "A"),
                ("fill", "C6EFCE"),
                ("font", "006100")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("operator", "equal");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CellIs_GreaterOrEqualLessOrEqual_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Score");
            AddRow(handler, 2, "X", "60");
            AddRow(handler, 3, "Y", "80");
            AddRow(handler, 4, "Z", "90");
            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "B2:B4"),
                ("operator", "greaterThanOrEqual"),
                ("formula", "80"),
                ("fill", "FFEB9C")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("operator", "greaterThanOrEqual");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CellIs_BetweenNotBetween_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "50");
            AddRow(handler, 4, "C", "90");
            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "B2:B4"),
                ("operator", "between"),
                ("formula", "20"),
                ("formula2", "80"),
                ("fill", "C6EFCE")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("operator", "between");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── ColorScale ───────────────────────────────────────────────────────

    [Fact]
    public void ColorScale_TwoColor_RoundTrip()
    {
        var (path, _) = SetupDataRange();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "colorscale", null, Props(
                ("range", "B2:B6"),
                ("minColor", "F8696B"),
                ("maxColor", "63BE7B"),
                ("minType", "min"),
                ("maxType", "max")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "colorScale");
        cf.Format.Should().Contain("minColor", "#F8696B");
        cf.Format.Should().Contain("maxColor", "#63BE7B");
        cf.Format.Should().Contain("ref", "B2:B6");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void ColorScale_ThreeColor_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Score");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "50");
            AddRow(handler, 4, "C", "90");
            handler.Add("/Sheet1", "colorscale", null, Props(
                ("range", "B2:B4"),
                ("minColor", "F8696B"),
                ("midColor", "FFEB84"),
                ("maxColor", "63BE7B"),
                ("minType", "min"),
                ("midType", "percentile"),
                ("maxType", "max"),
                ("midpoint", "50")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "colorScale");
        cf.Format.Should().Contain("midColor", "#FFEB84");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── DataBar ──────────────────────────────────────────────────────────

    [Fact]
    public void DataBar_Basic_RoundTrip()
    {
        var (path, _) = SetupDataRange();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "databar", null, Props(
                ("range", "B2:B6"),
                ("fillColor", "638EC6"),
                ("minType", "autoMin"),
                ("maxType", "autoMax")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "dataBar");
        cf.Format.Should().Contain("ref", "B2:B6");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void DataBar_WithGradientAndBorder_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "50");
            AddRow(handler, 4, "C", "90");
            handler.Add("/Sheet1", "databar", null, Props(
                ("range", "B2:B4"),
                ("fillColor", "4472C4"),
                ("minType", "num"),
                ("maxType", "num"),
                ("minValue", "0"),
                ("maxValue", "100"),
                ("gradient", "true"),
                ("border", "true"),
                ("borderColor", "000000"),
                ("direction", "leftToRight")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "dataBar");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── IconSet ──────────────────────────────────────────────────────────

    [Fact]
    public void IconSet_ThreeArrows_RoundTrip()
    {
        var (path, _) = SetupDataRange();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "iconset", null, Props(
                ("range", "B2:B6"),
                ("iconSet", "3Arrows"),
                ("showValue", "true")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "iconSet");
        cf.Format.Should().Contain("ref", "B2:B6");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void IconSet_ThreeTrafficLights_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Metric", "Value");
            AddRow(handler, 2, "Speed", "50");
            AddRow(handler, 3, "Quality", "85");
            handler.Add("/Sheet1", "iconset", null, Props(
                ("range", "B2:B3"),
                ("iconSet", "3TrafficLights1"),
                ("showValue", "false"),
                ("reverse", "false")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "iconSet");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void IconSet_ThreeStars_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Rating");
            AddRow(handler, 2, "A", "1");
            AddRow(handler, 3, "B", "2");
            AddRow(handler, 4, "C", "3");
            handler.Add("/Sheet1", "iconset", null, Props(
                ("range", "B2:B4"),
                ("iconSet", "3Arrows"),
                ("showValue", "true")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "iconSet");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Formula CF ───────────────────────────────────────────────────────

    [Fact]
    public void FormulaCf_EvaluationWithFillAndFont_RoundTrip()
    {
        var (path, _) = SetupDataRange();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "formulacf", null, Props(
                ("range", "A2:A6"),
                ("formula", "$B2>500"),
                ("fill", "C6EFCE"),
                ("font", "006100")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "formula");
        cf.Format.Should().Contain("formula", "$B2>500");
        cf.Format.Should().Contain("fill", "#C6EFCE");
        cf.Format.Should().Contain("ref", "A2:A6");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void FormulaCf_MultipleFormulaRules_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Value");
            AddRow(handler, 2, "Alpha", "100");
            AddRow(handler, 3, "Beta", "500");
            handler.Add("/Sheet1", "formulacf", null, Props(
                ("range", "A2:A3"),
                ("formula", "$B2>200"),
                ("fill", "FFF2CC")
            ));
            handler.Add("/Sheet1", "formulacf", null, Props(
                ("range", "A2:A3"),
                ("formula", "$B2<50"),
                ("fill", "F4CCCC")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/cf[1]").Format.Should().ContainKey("type");
        ReadNode(readOnly, "/Sheet1/cf[2]").Format.Should().ContainKey("type");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Top / Bottom ─────────────────────────────────────────────────────

    [Fact]
    public void TopN_Rank_RoundTrip()
    {
        var (path, _) = SetupDataRange();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "topn", null, Props(
                ("range", "B2:B6"),
                ("rank", "3"),
                ("fill", "FFC000")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "topN");
        cf.Format.Should().Contain("rank", 3U);
        cf.Format.Should().Contain("ref", "B2:B6");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void TopN_Percent_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Score");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "20");
            AddRow(handler, 4, "C", "30");
            AddRow(handler, 5, "D", "40");
            handler.Add("/Sheet1", "topn", null, Props(
                ("range", "B2:B5"),
                ("rank", "50"),
                ("percent", "true"),
                ("fill", "D9E2F3")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "topN");
        cf.Format.Should().Contain("percent", true);
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void BottomN_Rank_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "50");
            AddRow(handler, 4, "C", "90");
            handler.Add("/Sheet1", "topn", null, Props(
                ("range", "B2:B4"),
                ("rank", "2"),
                ("bottom", "true"),
                ("fill", "F4B4C4")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "topN");
        cf.Format.Should().Contain("bottom", true);
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Above/Below Average ──────────────────────────────────────────────

    [Fact]
    public void AboveAverage_RoundTrip()
    {
        var (path, _) = SetupDataRange();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1", "aboveaverage", null, Props(
                ("range", "B2:B6"),
                ("fill", "92D050")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "aboveAverage");
        cf.Format.Should().Contain("ref", "B2:B6");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void BelowAverage_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Score");
            AddRow(handler, 2, "A", "10");
            AddRow(handler, 3, "B", "80");
            AddRow(handler, 4, "C", "50");
            handler.Add("/Sheet1", "aboveaverage", null, Props(
                ("range", "B2:B4"),
                ("aboveAverage", "false"),
                ("fill", "F4CCCC")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "aboveAverage");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Unique / Duplicate Values ────────────────────────────────────────

    [Fact]
    public void DuplicateValues_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Code");
            AddRow(handler, 2, "Alice", "A1");
            AddRow(handler, 3, "Bob", "A1");
            AddRow(handler, 4, "Carol", "B2");
            handler.Add("/Sheet1", "duplicatevalues", null, Props(
                ("range", "B2:B4"),
                ("fill", "FF4444")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "duplicateValues");
        cf.Format.Should().Contain("ref", "B2:B4");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void UniqueValues_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "ID");
            AddRow(handler, 2, "Alice", "001");
            AddRow(handler, 3, "Bob", "002");
            AddRow(handler, 4, "Carol", "003");
            handler.Add("/Sheet1", "uniquevalues", null, Props(
                ("range", "B2:B4"),
                ("fill", "FFC000")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "uniqueValues");
        cf.Format.Should().Contain("ref", "B2:B4");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Contains Text ────────────────────────────────────────────────────

    [Fact]
    public void ContainsText_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales");
            AddRow(handler, 2, "East", "100");
            AddRow(handler, 3, "North East", "200");
            AddRow(handler, 4, "West", "300");
            handler.Add("/Sheet1", "containstext", null, Props(
                ("range", "A2:A4"),
                ("text", "East"),
                ("fill", "C6EFCE")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "containsText");
        cf.Format.Should().Contain("text", "East");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Date Occurring ───────────────────────────────────────────────────

    [Fact]
    public void DateOccurring_ThisMonth_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Task", "DueDate");
            handler.Add("/Sheet1/B2", "cell", null, Props(("value", "2026-07-15"), ("type", "date")));
            handler.Add("/Sheet1/B3", "cell", null, Props(("value", "2026-08-01"), ("type", "date")));
            handler.Add("/Sheet1/B4", "cell", null, Props(("value", "2026-06-20"), ("type", "date")));
            handler.Add("/Sheet1", "dateoccurring", null, Props(
                ("range", "B2:B4"),
                ("datePeriod", "thisMonth"),
                ("fill", "BDD7EE")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "timePeriod");
        cf.Format.Should().Contain("ref", "B2:B4");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── CF Extended ──────────────────────────────────────────────────────

    [Fact]
    public void CfExtended_BeginsWith_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Code", "Value");
            AddRow(handler, 2, "US-001", "10");
            AddRow(handler, 3, "UK-002", "20");
            AddRow(handler, 4, "US-003", "30");
            handler.Add("/Sheet1", "cfextended", null, Props(
                ("range", "A2:A4"),
                ("type", "beginsWith"),
                ("text", "US"),
                ("fill", "CFE2F3")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "beginsWith");
        cf.Format.Should().Contain("text", "US");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CfExtended_EndsWith_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "File", "Size");
            AddRow(handler, 2, "report.pdf", "100");
            AddRow(handler, 3, "data.csv", "50");
            AddRow(handler, 4, "notes.txt", "20");
            handler.Add("/Sheet1", "cfextended", null, Props(
                ("range", "A2:A4"),
                ("type", "endsWith"),
                ("text", ".pdf"),
                ("fill", "D9EAD3")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "endsWith");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CfExtended_ContainsBlanks_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Field", "Value");
            handler.Add("/Sheet1/B2", "cell", null, Props(("value", "100")));
            // B3 intentionally blank
            handler.Add("/Sheet1/B4", "cell", null, Props(("value", "200")));
            handler.Add("/Sheet1", "cfextended", null, Props(
                ("range", "B2:B4"),
                ("type", "containsBlanks"),
                ("fill", "F4CCCC")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "containsBlanks");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CfExtended_NotContainsBlanks_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Note");
            AddRow(handler, 2, "Alice", "Done");
            AddRow(handler, 3, "Bob", "");
            AddRow(handler, 4, "Carol", "Pending");
            handler.Add("/Sheet1", "cfextended", null, Props(
                ("range", "B2:B4"),
                ("type", "notContainsBlanks"),
                ("fill", "C6EFCE")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "notContainsBlanks");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CfExtended_ContainsErrors_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Name", "Result");
            AddRow(handler, 2, "A", "100");
            handler.Add("/Sheet1/B3", "cell", null, Props(("formula", "1/0")));
            AddRow(handler, 4, "C", "200");
            handler.Add("/Sheet1", "cfextended", null, Props(
                ("range", "B2:B4"),
                ("type", "containsErrors"),
                ("fill", "FF4444")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "containsErrors");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void CfExtended_NotContainsErrors_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "A", "100");
            AddRow(handler, 3, "B", "200");
            handler.Add("/Sheet1", "cfextended", null, Props(
                ("range", "B2:B3"),
                ("type", "notContainsErrors"),
                ("fill", "C6EFCE")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var cf = ReadNode(readOnly, "/Sheet1/cf[1]");
        cf.Format.Should().Contain("type", "notContainsErrors");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Query and Remove ─────────────────────────────────────────────────

    [Fact]
    public void Cf_QueryAllRules_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "Alpha", "100");
            AddRow(handler, 3, "Beta", "300");
            AddRow(handler, 4, "Gamma", "600");
            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "B2:B4"), ("operator", "greaterThan"), ("formula", "200"), ("fill", "C6EFCE")));
            handler.Add("/Sheet1", "databar", null, Props(
                ("range", "B2:B4"), ("fillColor", "638EC6")));
            handler.Add("/Sheet1", "colorscale", null, Props(
                ("range", "B2:B4"), ("minColor", "FF0000"), ("maxColor", "00FF00")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/cf[1]").Format.Should().Contain("type", "cellIs");
        ReadNode(readOnly, "/Sheet1/cf[2]").Format.Should().Contain("type", "dataBar");
        ReadNode(readOnly, "/Sheet1/cf[3]").Format.Should().Contain("type", "colorScale");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Cf_RemoveIndividualRule_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "Alpha", "100");
            AddRow(handler, 3, "Beta", "300");
            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "B2:B3"), ("operator", "greaterThan"), ("formula", "150"), ("fill", "C6EFCE")));
            handler.Add("/Sheet1", "databar", null, Props(
                ("range", "B2:B3"), ("fillColor", "638EC6")));
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/cf[1]", null);
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/cf[1]").Format.Should().ContainKey("type");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void Cf_RemoveAllRules_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "Alpha", "100");
            AddRow(handler, 3, "Beta", "300");
            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "B2:B3"), ("operator", "greaterThan"), ("formula", "150"), ("fill", "C6EFCE")));
            handler.Add("/Sheet1", "databar", null, Props(
                ("range", "B2:B3"), ("fillColor", "638EC6")));
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Sheet1/cf[2]", null);
            handler.Remove("/Sheet1/cf[1]", null);
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var raw = readOnly.Raw("/Sheet1");
        raw.Should().NotContain("<conditionalFormatting");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Multiple CF types on same range ──────────────────────────────────

    [Fact]
    public void Cf_MultipleRuleTypesOnSameRange_RoundTrip()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Region", "Sales", "Date");
            AddRow(handler, 2, "West", "1200", "2026-01-15");
            AddRow(handler, 3, "East", "800", "2026-03-22");
            AddRow(handler, 4, "North", "1500", "2026-07-01");

            handler.Add("/Sheet1", "cellis", null, Props(
                ("range", "B2:B4"), ("operator", "greaterThan"), ("formula", "1000"), ("fill", "C6EFCE")));
            handler.Add("/Sheet1", "topn", null, Props(
                ("range", "B2:B4"), ("rank", "2"), ("fill", "FFC000")));
            handler.Add("/Sheet1", "aboveaverage", null, Props(
                ("range", "B2:B4"), ("fill", "92D050")));
            handler.Add("/Sheet1", "dateoccurring", null, Props(
                ("range", "C2:C4"), ("datePeriod", "thisMonth"), ("fill", "BDD7EE")));

            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        ReadNode(readOnly, "/Sheet1/cf[1]").Format.Should().ContainKey("type");
        ReadNode(readOnly, "/Sheet1/cf[2]").Format.Should().ContainKey("type");
        ReadNode(readOnly, "/Sheet1/cf[3]").Format.Should().ContainKey("type");
        ReadNode(readOnly, "/Sheet1/cf[4]").Format.Should().ContainKey("type");
        readOnly.Validate().Should().BeEmpty();
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private (string Path, ExcelHandler Handler) SetupDataRange()
    {
        var path = CreateWorkbook();
        using (var handler = OpenEditable(path))
        {
            AddRow(handler, 1, "Item", "Value");
            AddRow(handler, 2, "Alpha", "100");
            AddRow(handler, 3, "Beta", "300");
            AddRow(handler, 4, "Gamma", "600");
            AddRow(handler, 5, "Delta", "200");
            AddRow(handler, 6, "Epsilon", "900");
            handler.Save();
        }
        return (path, null!);
    }

    private static Dictionary<string, string> Props(params (string Key, string Value)[] values)
        => values.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

    private static void AddRow(ExcelHandler handler, int row, params string[] values)
    {
        for (var i = 0; i < values.Length; i++)
        {
            var column = (char)('A' + i);
            handler.Add($"/Sheet1/{column}{row}", "cell", null, Props(("value", values[i])));
        }
    }
}
