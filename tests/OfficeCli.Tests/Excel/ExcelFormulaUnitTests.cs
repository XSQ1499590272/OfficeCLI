// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelFormulaUnitTests
{
    public static IEnumerable<object?[]> FormulaDispatchCases()
    {
        yield return Case.Number("arithmetic", "SUMPRODUCT(A1:A3,B1:B3)", 140);
        yield return Case.Number("references", "OFFSET(B1,1,0)", 20);
        yield return Case.Number("lookup", "XLOOKUP(2,A1:A3,B1:B3)", 20);
        yield return Case.Text("text", "TEXTAFTER(\"OfficeCLI-Excel\",\"-\")", "Excel");
        yield return Case.Number("date/time", "YEAR(DATE(2026,7,8))", 2026);
        yield return Case.Number("dynamic array/spill", "SEQUENCE(2,2)", 1);
        yield return Case.Number("financial", "PMT(0.05/12,12,1000)", -85.6074817884674, 1e-9);
        yield return Case.Number("statistics", "STDEV.S(A1:A3)", 1);
        yield return Case.Number("regression", "SLOPE(B1:B3,A1:A3)", 10);
        yield return Case.Number("securities", "COUPNUM(DATE(2026,1,1),DATE(2027,1,1),2,0)", 2);
        yield return Case.Number("special functions", "ERF(1)", 0.842700792949715, 1e-12);
        yield return Case.Number("error functions", "ERROR.TYPE(NA())", 7);
    }

    [Theory]
    [MemberData(nameof(FormulaDispatchCases))]
    public void FormulaEvaluator_CoversAllImplementedDispatchFamilies(
        string family,
        string formula,
        string expectedStatus,
        double? expectedNumber,
        string? expectedText,
        string? expectedError,
        double tolerance)
    {
        var report = CreateEvaluator().EvaluateForReport(formula);

        report.Status.ToString().Should().Be(expectedStatus, family);
        report.Result.Should().NotBeNull(family);
        report.Result!.ErrorValue.Should().Be(expectedError, family);

        if (expectedNumber.HasValue)
            report.Result.NumericValue.Should().BeApproximately(expectedNumber.Value, tolerance, family);

        if (expectedText is not null)
            report.Result.AsString().Should().Be(expectedText, family);
    }

    [Fact]
    public void FormulaEvaluator_PropagatesExcelErrorsAndUnsupportedFunctions()
    {
        var evaluator = CreateEvaluator();
        using var workbook = CreateWorkbook(("Sheet1", Sheet(Num("A1", 1))));

        evaluator.EvaluateForReport("MOD(1,0)").Should().Match<EvalReport>(
            report => report.Status == EvalReportStatus.Error && report.Result!.ErrorValue == "#DIV/0!");
        evaluator.EvaluateForReport("SQRT(-1)").Should().Match<EvalReport>(
            report => report.Status == EvalReportStatus.Error && report.Result!.ErrorValue == "#NUM!");
        evaluator.EvaluateForReport("VALUE(\"abc\")").Should().Match<EvalReport>(
            report => report.Status == EvalReportStatus.Error && report.Result!.ErrorValue == "#VALUE!");
        new FormulaEvaluator(GetSheetData(workbook, "Sheet1"), workbook.WorkbookPart)
            .EvaluateForReport("MissingSheet!A1")
            .Should().Match<EvalReport>(
                report => report.Status == EvalReportStatus.Error && report.Result!.ErrorValue == "#REF!");
        evaluator.EvaluateForReport("TRIMMEAN(A1:A3)").Status.Should().Be(EvalReportStatus.NotEvaluated);
        evaluator.EvaluateForReport("NO_SUCH_FN(1)").Status.Should().Be(EvalReportStatus.NotEvaluated);
        evaluator.EvaluateForReport("A10+1").Should().Match<EvalReport>(
            report => report.Status == EvalReportStatus.Evaluated && report.Result!.NumericValue == 1);
    }

    [Fact]
    public void ModernFunctionQualifier_QualifiesDynamicArraysAndQuotesSheetRefs()
    {
        ModernFunctionQualifier.IsDynamicArrayFormula("SORT(A1:A3)").Should().BeTrue();
        ModernFunctionQualifier.IsDynamicArrayFormula("\"SORT(A1:A3)\"&A1").Should().BeFalse();

        ModernFunctionQualifier.Qualify("FILTER(A1:A3,B1:B3>0)+SEQUENCE(2)")
            .Should().Be("_xlfn._xlws.FILTER(A1:A3,B1:B3>0)+_xlfn.SEQUENCE(2)");
        ModernFunctionQualifier.Unqualify("_xlfn._xlws.FILTER(A1:A3,B1:B3>0)+_xlfn.SEQUENCE(2)")
            .Should().Be("FILTER(A1:A3,B1:B3>0)+SEQUENCE(2)");
        ModernFunctionQualifier.AutoQuoteSheetRefs("SUM(My Sheet!A1,2026.Q1!B2,Plain_1!C3,\"My Sheet!D4\")")
            .Should().Be("SUM('My Sheet'!A1,'2026.Q1'!B2,Plain_1!C3,\"My Sheet!D4\")");
        ModernFunctionQualifier.QuoteSheetNameForRef("O'Brien Sales")
            .Should().Be("'O''Brien Sales'");
    }

    [Fact]
    public void FormulaRefShifter_RewritesRowsColumnsCopiesAndSheetRenames()
    {
        FormulaRefShifter.Shift(
                "SUM(A1,$B$2,B3:C4,\"B5\",Table1[Col],'Other Sheet'!B2)",
                "Sheet1",
                "Sheet1",
                FormulaShiftDirection.ColumnsRight,
                2)
            .Should().Be("SUM(A1,$C$2,C3:D4,\"B5\",Table1[Col],'Other Sheet'!B2)");

        FormulaRefShifter.Shift("SUM(B2:B4)", "Sheet1", "Sheet1", FormulaShiftDirection.RowsUp, 3)
            .Should().Be("SUM(B2:B3)");
        FormulaRefShifter.Shift("B3", "Sheet1", "Sheet1", FormulaShiftDirection.RowsUp, 3)
            .Should().Be("#REF!");

        FormulaRefShifter.ApplyRowRenumberMap("SUM(A2:A4,'My Sheet'!B2:B4)", "Sheet1", "Sheet1", new Dictionary<int, int>
        {
            [2] = 3,
            [3] = 4,
            [4] = 5
        }).Should().Be("SUM(A3:A5,'My Sheet'!B2:B4)");

        FormulaRefShifter.ApplyColRenumberMap("SUM(B1:C1,'My Sheet'!B1:C1)", "Sheet1", "My Sheet", new Dictionary<int, int>
        {
            [2] = 4,
            [3] = 5
        }).Should().Be("SUM(B1:C1,'My Sheet'!D1:E1)");

        FormulaRefShifter.ApplyCopyDelta("SUM(A1,$B2,C$3,\"A1\",Table1[A1])", "Sheet1", "Sheet1", 2, 1)
            .Should().Be("SUM(C2,$B3,E$3,\"A1\",Table1[A1])");

        FormulaRefShifter.RenameSheetRef("SUM('Old Name'!A1,Summary!B2,\"Old Name!C3\")", "'Old Name'!", "'New Name'!")
            .Should().Be("SUM('New Name'!A1,Summary!B2,\"Old Name!C3\")");
    }

    [Fact]
    public void FormulaCacheHelpers_ClassifyAllowlistAndComputedAgreement()
    {
        ExcelHandler.FormulaFunctionsCovered(
                "_xlfn._xlws.FILTER(A1:A3,B1:B3>0)+SUM(C1:C3)",
                new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "FILTER", "SUM" })
            .Should().BeTrue();
        ExcelHandler.FormulaFunctionsCovered(
                "_xlfn._xlws.FILTER(A1:A3,B1:B3>0)+SUM(C1:C3)",
                new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "SUM" })
            .Should().BeFalse();
        ExcelHandler.FormulaFunctionsCovered(
                "A1",
                new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "SUM" })
            .Should().BeFalse();

        ExcelHandler.CachedComputedAgree("1.0000000001", "1.0000000002").Should().BeTrue();
        ExcelHandler.CachedComputedAgree("1.0", "1.01").Should().BeFalse();
        ExcelHandler.CachedComputedAgree("hello", "world").Should().BeFalse();
        ExcelHandler.CachedComputedAgree("#REF!", "#VALUE!").Should().BeFalse();
    }

    private static FormulaEvaluator CreateEvaluator() => new(Sheet(
        Num("A1", 1), Num("B1", 10),
        Num("A2", 2), Num("B2", 20),
        Num("A3", 3), Num("B3", 30),
        Str("C1", "dup"), Str("C2", "dup"), Str("C3", "tail")));

    private static SheetData Sheet(params Cell[] cells)
    {
        var rows = cells
            .GroupBy(cell => new string(cell.CellReference!.Value!.Where(char.IsDigit).ToArray()))
            .Select(group => new Row(group.OrderBy(cell => cell.CellReference!.Value).Select(cell => (OpenXmlElement)cell).ToArray())
            {
                RowIndex = uint.Parse(group.Key)
            });
        return new SheetData(rows);
    }

    private static Cell Num(string reference, double value) => new()
    {
        CellReference = reference,
        CellValue = new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture))
    };

    private static Cell Str(string reference, string value) => new()
    {
        CellReference = reference,
        CellValue = new CellValue(value),
        DataType = CellValues.String
    };

    private static SpreadsheetDocument CreateWorkbook(params (string Name, SheetData Data)[] sheets)
    {
        var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var sheetsNode = workbookPart.Workbook.AppendChild(new Sheets());
        uint sheetId = 1;
        foreach (var (name, data) in sheets)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet((SheetData)data.CloneNode(true));
            sheetsNode.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = name
            });
        }

        workbookPart.Workbook.Save();
        return document;
    }

    private static SheetData GetSheetData(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var workbook = workbookPart.Workbook!;
        var sheet = workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .Single(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        return worksheetPart.Worksheet!.GetFirstChild<SheetData>()!;
    }

    private static object?[] Pack(
        string family,
        string expectedStatus,
        string formula,
        double? expectedNumber = null,
        string? expectedText = null,
        string? expectedError = null,
        double tolerance = 1e-12) =>
        [family, formula, expectedStatus, (object?)expectedNumber, expectedText, expectedError, tolerance];

    private static class Case
    {
        public static object?[] Number(string family, string formula, double value, double tolerance = 1e-12) =>
            Pack(family, nameof(EvalReportStatus.Evaluated), formula, expectedNumber: value, tolerance: tolerance);

        public static object?[] Text(string family, string formula, string value) =>
            Pack(family, nameof(EvalReportStatus.Evaluated), formula, expectedText: value);
    }
}
