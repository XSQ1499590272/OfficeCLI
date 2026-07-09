// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelUnitTests
{
    [Fact]
    public void FormulaEvaluator_EvaluatesRepresentativeFamilies()
    {
        var evaluator = new FormulaEvaluator(Sheet(
            Num("A1", 1), Num("B1", 10),
            Num("A2", 2), Num("B2", 20),
            Num("A3", 3), Num("B3", 30)));

        evaluator.TryEvaluateFull("SUM(A1:A3)")!.NumericValue.Should().Be(6);
        evaluator.TryEvaluateFull("VLOOKUP(2,A1:B3,2,FALSE)")!.NumericValue.Should().Be(20);
        evaluator.TryEvaluateFull("DATE(2026,7,8)")!.NumericValue.Should().Be(new DateTime(2026, 7, 8).ToOADate());
        evaluator.TryEvaluateFull("SEQUENCE(2,2)")!.NumericValue.Should().Be(1);
        evaluator.TryEvaluateFull("PMT(0.05/12,12,1000)")!.NumericValue.Should().BeApproximately(-85.607, 0.001);
        evaluator.TryEvaluateFull("STDEV.S(A1:A3)")!.NumericValue.Should().Be(1);
        evaluator.TryEvaluateFull("COUPNUM(DATE(2026,1,1),DATE(2027,1,1),2,0)")!.NumericValue.Should().Be(2);
    }

    [Fact]
    public void FormulaEvaluator_ReturnsExcelErrorsForInvalidMath()
    {
        var evaluator = new FormulaEvaluator(new SheetData());

        evaluator.EvaluateForReport("SQRT(-1)").Should().Match<EvalReport>(
            r => r.Status == EvalReportStatus.Error && r.Result!.ErrorValue == "#NUM!");
        evaluator.EvaluateForReport("MOD(1,0)").Should().Match<EvalReport>(
            r => r.Status == EvalReportStatus.Error && r.Result!.ErrorValue == "#DIV/0!");
        evaluator.EvaluateForReport("NO_SUCH_FN(1)").Status.Should().Be(EvalReportStatus.NotEvaluated);
    }

    [Fact]
    public void ExcelDataFormatter_FormatsDatesPercentAndBuiltIns()
    {
        ExcelDataFormatter.ResolveBuiltInFormatCode(14).Should().Be("m/d/yy");
        ExcelDataFormatter.ResolveBuiltInFormatCode(49).Should().Be("@");
        ExcelDataFormatter.TryFormat(0.25, 9, null).Should().Be("25%");
        ExcelDataFormatter.TryFormat(new DateTime(2026, 7, 8).ToOADate(), 14, null)
            .Should().Contain("26");
    }

    [Fact]
    public void AttributeFilter_ParsesAndAppliesExcelCellSelectors()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/Sheet1/A1",
                Type = "cell",
                Text = "Revenue Jan",
                Format = new() { ["value"] = "42", ["font.bold"] = true, ["width"] = "12pt" }
            },
            new()
            {
                Path = "/Sheet1/A2",
                Type = "cell",
                Text = "Cost Feb",
                Format = new() { ["value"] = "7" }
            }
        };

        AttributeFilter.Apply(nodes, AttributeFilter.Parse("cell[value>=40][font.bold]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A1");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("cell[text~=r\"Revenue\\s+Jan\"]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A1");
        AttributeFilter.ApplyExpr(nodes, AttributeFilter.ParseExpr("cell[value=7 or font.bold]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A1", "/Sheet1/A2");
    }

    [Fact]
    public void ExcelHelperValidation_RejectsBadSheetNamesRangesAndCellValues()
    {
        FluentActions.Invoking(() => ExcelHandler.ValidateSheetName("Data"))
            .Should().NotThrow();
        FluentActions.Invoking(() => ExcelHandler.ValidateSheetName("Bad/Name"))
            .Should().Throw<ArgumentException>();
        ExcelHandler.ValidateSqref("A1 B2:C5 A:A 1:3", "range").Should().Be("A1 B2:C5 A:A 1:3");
        FluentActions.Invoking(() => ExcelHandler.ValidateSqref("not-a-range", "range"))
            .Should().Throw<ArgumentException>();
        FluentActions.Invoking(() => ExcelHandler.ValidateSparklineRange("not a range"))
            .Should().Throw<ArgumentException>();
        FluentActions.Invoking(() => ExcelHandler.EnsureCellValueLength(new string('x', ExcelHandler.MaxCellTextLength + 1), "A1"))
            .Should().Throw<ArgumentException>();
    }

    [Fact]
    public void HtmlScreenshot_AutoGridColumns_IsStable()
    {
        HtmlScreenshot.AutoGridColumns(1, 160, 90).Should().Be(1);
        HtmlScreenshot.AutoGridColumns(2, 160, 90).Should().Be(1);
        HtmlScreenshot.AutoGridColumns(4, 100, 100).Should().Be(2);
        HtmlScreenshot.AutoGridColumns(9, 100, 100).Should().Be(3);
    }

    private static SheetData Sheet(params Cell[] cells)
    {
        var rows = cells
            .GroupBy(c => new string(c.CellReference!.Value!.Where(char.IsDigit).ToArray()))
            .Select(g => new Row(g.OrderBy(c => c.CellReference!.Value).Select(c => (OpenXmlElement)c).ToArray())
            {
                RowIndex = uint.Parse(g.Key)
            });
        return new SheetData(rows);
    }

    private static Cell Num(string reference, double value) => new()
    {
        CellReference = reference,
        CellValue = new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture))
    };
}
