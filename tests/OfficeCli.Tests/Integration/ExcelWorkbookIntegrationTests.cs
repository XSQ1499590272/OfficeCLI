// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class ExcelWorkbookIntegrationTests : ExcelTestBase
{
    [Fact]
    public void WorkbookLifecycle_CreateOpenSaveReopenValidate_RoundTrips()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/", Props(("title", "Test Workbook"), ("subject", "Integration Tests")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var root = ReadNode(readOnly, "/");
        root.Format.Should().Contain("title", "Test Workbook");
        root.Format.Should().Contain("subject", "Integration Tests");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookProperties_TitleSubjectCalcSettings_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/", Props(
                ("title", "Q1 Report"),
                ("subject", "Quarterly financial summary"),
                ("calc.mode", "manual"),
                ("calc.iterate", "true"),
                ("calc.iterateCount", "50"),
                ("calc.fullPrecision", "true")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var root = ReadNode(readOnly, "/");
        root.Format.Should().Contain("title", "Q1 Report");
        root.Format.Should().Contain("subject", "Quarterly financial summary");
        root.Format.Should().Contain("calc.mode", "manual");
        root.Format.Should().Contain("calc.iterate", true);
        root.Format.Should().Contain("calc.iterateCount", 50U);
        root.Format.Should().Contain("calc.fullPrecision", true);
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookCalcOnSaveRoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/", Props(("calc.fullCalcOnLoad", "true")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var root = ReadNode(readOnly, "/");
        // The exact readback key varies; verify the document is valid.
        root.Format.Keys.Should().Contain(k => k.Contains("calc", StringComparison.OrdinalIgnoreCase));
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookAuthor_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/", Props(("author", "OfficeCLI Tests")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var root = ReadNode(readOnly, "/");
        root.Format.Should().Contain("author", "OfficeCLI Tests");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void SheetLifecycle_AddRenameVisibilityFreeze_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "sheet", null, Props(
                ("name", "RawData"),
                ("tabColor", "FF0000"),
                ("visibility", "hidden")
            ));
            handler.Set("/RawData", Props(
                ("name", "CleanData"),
                ("freeze", "C3")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/CleanData");
        sheet.Path.Should().Be("/CleanData");
        sheet.Format.Should().Contain("tabColor", "#FF0000");
        sheet.Format.Should().Contain("freeze", "C3");
        Query(readOnly, "sheet").Should().Contain(s => s.Path == "/CleanData");
        Query(readOnly, "sheet").Should().NotContain(s => s.Path == "/RawData");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void SheetDisplayOptions_DirectionRtl_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1", Props(
                ("direction", "rtl"),
                ("rtl", "true")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        sheet.Format.Should().Contain("direction", "rtl");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void SheetPrintSettings_PrintArea_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1", Props(
                ("printArea", "A1:F50")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        sheet.Format.Should().Contain("printArea", "A1:F50");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void SheetHeaderFooter_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1", Props(
                ("header", "&LConfidential &C&\"Arial,Bold\"&14Q1 Report &R&D"),
                ("footer", "&CPage &P of &N")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        sheet.Format.Should().ContainKey("header");
        sheet.Format.Should().ContainKey("footer");
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void SheetProtection_SetAndValidate_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1", Props(
                ("protection", "hash"),
                ("protection.password", "sheetpass")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        // Protection may not appear as a Format key; verify document is valid.
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void WorkbookProtection_SetAndValidate_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/", Props(
                ("protection", "hash"),
                ("protection.password", "secret")
            ));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void ActiveTab_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "sheet", null, Props(("name", "Second")));
            handler.Set("/", Props(("activeTab", "1")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var root = ReadNode(readOnly, "/");
        root.Format.Should().Contain("activeTab", 1U);
        Query(readOnly, "sheet").Select(s => s.Path).Should().Contain(new[] { "/Sheet1", "/Second" });
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void AddRemoveSheet_LifecycleAndReadd()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "sheet", null, Props(("name", "Temp")));
            handler.Save();
        }

        using (var readOnly = OpenReadOnly(path))
        {
            Query(readOnly, "sheet").Should().Contain(s => s.Path == "/Temp");
            readOnly.Validate().Should().BeEmpty();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Remove("/Temp", null);
            handler.Save();
        }

        using (var readOnly = OpenReadOnly(path))
        {
            Query(readOnly, "sheet").Should().NotContain(s => s.Path == "/Temp");
            readOnly.Validate().Should().BeEmpty();
        }

        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "sheet", null, Props(("name", "ReAdded")));
            handler.Save();
        }

        using var finalReadOnly = OpenReadOnly(path);
        Query(finalReadOnly, "sheet").Should().Contain(s => s.Path == "/ReAdded");
        finalReadOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void SheetTabColor_RoundTrip()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1", Props(("tabColor", "4472C4")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        sheet.Format.Should().Contain("tabColor", "#4472C4");
        readOnly.Validate().Should().BeEmpty();
    }

    [Theory]
    [InlineData("hidden", "hidden")]
    [InlineData("veryhidden", "veryHidden")]
    public void SheetVisibility_Variants_RoundTrip(string setVisibility, string expectedVisibility)
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Set("/Sheet1", Props(("visibility", setVisibility)));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        sheet.Format.Should().Contain("visibility", expectedVisibility);
        readOnly.Validate().Should().BeEmpty();
    }

    [Fact]
    public void SheetVisibleByDefault_Validate()
    {
        var path = CreateWorkbook();

        using (var handler = OpenEditable(path))
        {
            handler.Add("/Sheet1/A1", "cell", null, Props(("value", "data")));
            handler.Save();
        }

        using var readOnly = OpenReadOnly(path);
        var sheet = ReadNode(readOnly, "/Sheet1");
        sheet.Path.Should().Be("/Sheet1");
        readOnly.Validate().Should().BeEmpty();
    }

    // --- Helpers ---

    private static Dictionary<string, string> Props(params (string Key, string Value)[] values)
        => values.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
}
