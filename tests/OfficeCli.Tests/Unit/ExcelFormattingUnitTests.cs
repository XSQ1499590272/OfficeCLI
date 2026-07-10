// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Handlers;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class ExcelFormattingUnitTests
{
    public static IEnumerable<object[]> BuiltInFormatCases()
    {
        yield return ["general", 0u, "General"];
        yield return ["integer", 1u, "0"];
        yield return ["decimal", 2u, "0.00"];
        yield return ["percent", 9u, "0%"];
        yield return ["scientific", 11u, "0.00E+00"];
        yield return ["fraction", 12u, "# ?/?"];
        yield return ["currency/accounting", 39u, "#,##0.00;(#,##0.00)"];
        yield return ["date", 14u, "m/d/yy"];
        yield return ["time", 20u, "h:mm"];
        yield return ["text", 49u, "@"];
    }

    public static IEnumerable<object[]> RawReadbackFamilies()
    {
        yield return new object[] { "general", 12.5, 0u, null! };
        yield return new object[] { "integer", 12, 1u, null! };
        yield return new object[] { "decimal", 12.34, 2u, null! };
        yield return new object[] { "scientific", 12.34, 11u, null! };
        yield return new object[] { "fraction", 1.25, 12u, null! };
        yield return new object[] { "currency/accounting", -42, 39u, null! };
        yield return new object[] { "text", 42, 49u, null! };
        yield return ["custom sections without percent/date semantics", 42, 164u, "[Blue]0;[Red](0);0;[Magenta]@"];
    }

    [Theory]
    [MemberData(nameof(BuiltInFormatCases))]
    public void ResolveBuiltInFormatCode_CoversExcelReadbackFamilies(string family, uint numFmtId, string expectedCode)
    {
        ExcelDataFormatter.ResolveBuiltInFormatCode(numFmtId).Should().Be(expectedCode, family);
    }

    [Theory]
    [MemberData(nameof(RawReadbackFamilies))]
    public void TryFormat_LeavesNonDateAndNonPercentFamiliesOnRawValuePath(
        string family,
        double value,
        uint numFmtId,
        string? customFormatCode)
    {
        ExcelDataFormatter.TryFormat(value, numFmtId, customFormatCode).Should().BeNull(family);
    }

    [Fact]
    public void TryFormat_FormatsPercentBuiltInsAndCustomPositiveNegativeZeroSections()
    {
        ExcelDataFormatter.TryFormat(0.256, 9, null).Should().Be("26%");
        ExcelDataFormatter.TryFormat(0.256, 10, null).Should().Be("25.60%");

        const string custom = "0%;[Red](0%);zero 0%;[Blue]@";
        ExcelDataFormatter.TryFormat(0.42, 164, custom).Should().Be("42%");
        ExcelDataFormatter.TryFormat(-0.42, 164, custom).Should().Be("(42%)");
        ExcelDataFormatter.TryFormat(0, 164, custom).Should().Be("zero 0%");
    }

    [Fact]
    public void TryFormat_FormatsDateTimeAnd1904Epoch()
    {
        var day = new DateTime(2026, 7, 8, 0, 0, 0, DateTimeKind.Unspecified).ToOADate();
        var minute = new DateTime(2026, 7, 8, 14, 30, 0, DateTimeKind.Unspecified).ToOADate();
        var second = new DateTime(2026, 7, 8, 14, 30, 45, DateTimeKind.Unspecified).ToOADate();

        ExcelDataFormatter.TryFormat(day, 14, null).Should().Be("2026-07-08");
        ExcelDataFormatter.TryFormat(minute, 20, null).Should().Be("2026-07-08 14:30");
        ExcelDataFormatter.TryFormat(second, 21, null).Should().Be("2026-07-08 14:30:45");
        ExcelDataFormatter.TryFormat(0, 14, null, date1904: true).Should().Be("1904-01-01");
    }

    [Fact]
    public void GetCellFormat_ReturnsBuiltInAndCustomCodesFromStylesheet()
    {
        using var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        workbookPart.AddNewPart<WorksheetPart>().Worksheet = new Worksheet(new SheetData());
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new NumberingFormats(
                new NumberingFormat
                {
                    NumberFormatId = 164u,
                    FormatCode = "[Blue]0;[Red](0);0;[Magenta]@"
                })
            {
                Count = 1
            },
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(
                new CellFormat(),
                new CellFormat { NumberFormatId = 14u },
                new CellFormat { NumberFormatId = 164u })
            {
                Count = 3
            });
        stylesPart.Stylesheet.Save();

        var builtInCell = new Cell { StyleIndex = 1u };
        var customCell = new Cell { StyleIndex = 2u };

        ExcelDataFormatter.GetCellFormat(builtInCell, workbookPart).Should().Be((14u, (string?)null));
        ExcelDataFormatter.GetCellFormat(customCell, workbookPart)
            .Should().Be((164u, "[Blue]0;[Red](0);0;[Magenta]@"));
    }
}
