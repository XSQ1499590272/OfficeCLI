// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using P = DocumentFormat.OpenXml.Presentation;

using OfficeCli.Tests.Excel;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class ExcelSelectorsAndRawUnitTests
{
    private static readonly XNamespace SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Fact]
    public void AttributeFilter_ParseAndApply_CoversSelectorOperators()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/Sheet1/A1",
                Type = "cell",
                Text = "Revenue Jan",
                Format = new()
                {
                    ["value"] = "42",
                    ["font.bold"] = true,
                    ["font.color"] = "#FF0000",
                    ["width"] = "12pt"
                }
            },
            new()
            {
                Path = "/Sheet1/A2",
                Type = "cell",
                Text = "Cost Feb",
                Format = new()
                {
                    ["value"] = "7",
                    ["width"] = "8pt"
                }
            }
        };

        AttributeFilter.Apply(nodes, AttributeFilter.Parse("cell[value=42]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A1");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("cell[value!=42]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A2");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("cell[text~=Revenue]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A1");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("cell[text~=r\"Feb$\"]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A2");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("cell[value>10]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A1");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("cell[font.bold]"))
            .Select(n => n.Path).Should().Equal("/Sheet1/A1");
    }

    [Fact]
    public void AttributeFilter_ExpressionTree_SupportsAndOrAndCellAliasNormalization()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/Sheet1/A1",
                Type = "cell",
                Text = "Revenue Jan",
                Format = new() { ["font.bold"] = true, ["value"] = "42" }
            },
            new()
            {
                Path = "/Sheet1/A2",
                Type = "cell",
                Text = "Cost Feb",
                Format = new() { ["value"] = "7" }
            }
        };

        var expr = AttributeFilter.ParseExpr("Sheet1!cell[bold=true and value>40 or text~=r\"Feb$\"]");
        expr.Should().NotBeNull();

        var normalized = AttributeFilter.NormalizeKeysExpr(expr!, ExcelHandler.ResolveCellAttributeAlias);
        AttributeFilter.ApplyExpr(nodes, normalized)
            .Select(n => n.Path)
            .Should().Equal("/Sheet1/A1", "/Sheet1/A2");
    }

    [Fact]
    public void AttributeFilter_ReportsUnknownKeysAndNonNumericComparisons()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/Sheet1/A1",
                Type = "cell",
                Text = "42",
                Format = new()
                {
                    ["font.bold"] = true,
                    ["width"] = "12pt"
                }
            }
        };

        var expr = AttributeFilter.ParseExpr("cell[blod=true and width>wide]");
        var (results, warnings) = AttributeFilter.ApplyExprWithWarnings(nodes, expr);

        results.Should().BeEmpty();
        warnings.Should().ContainSingle(w => w.Kind == "unknown_key" && w.Key == "blod");
        warnings.Should().ContainSingle(w => w.Kind == "non_numeric" && w.Value == "wide");
        warnings.Single(w => w.Kind == "unknown_key").Available.Should().Contain("font.bold");
    }

    [Fact]
    public void RawXmlHelper_Execute_CoversCoreMutationActionsAndNamespacePreservation()
    {
        var worksheet = new Worksheet(new SheetData(
            new Row { RowIndex = 2u },
            new Row { RowIndex = 4u }));

        RawXmlHelper.Execute(worksheet, "/x:worksheet/x:sheetData", "prepend", "<row r=\"1\"/>").Should().Be(1);
        RawXmlHelper.Execute(worksheet, "/x:worksheet/x:sheetData", "append", "<row r=\"5\"/>").Should().Be(1);
        RawXmlHelper.Execute(worksheet, "//x:row[@r='4']", "insertbefore", "<row r=\"3\"/>").Should().Be(1);
        RawXmlHelper.Execute(worksheet, "//x:row[@r='5']", "insertafter", "<row r=\"6\"/>").Should().Be(1);
        RawXmlHelper.Execute(worksheet, "//x:row[@r='3']", "replace", "<row r=\"3\" customAttr=\"keep\"/>").Should().Be(1);
        RawXmlHelper.Execute(worksheet, "//x:row[@r='6']", "remove", null).Should().Be(1);
        RawXmlHelper.Execute(worksheet, "//x:row[@r='5']", "setattr", "spans=1:5").Should().Be(1);

        var rows = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();
        rows.Select(r => r.RowIndex!.Value).Should().Equal(1u, 2u, 3u, 4u, 5u);
        rows.Single(r => r.RowIndex!.Value == 3u).GetAttribute("customAttr", "").Value.Should().Be("keep");
        rows.Single(r => r.RowIndex!.Value == 5u).Spans!.InnerText.Should().Be("1:5");

        var xdoc = XDocument.Parse(worksheet.OuterXml);
        xdoc.Root!.Name.Namespace.Should().Be(SpreadsheetNs);
        xdoc.Descendants(SpreadsheetNs + "row").Should().HaveCount(5);
    }

    [Fact]
    public void RawXmlHelper_Execute_ReordersKnownSchemaChildrenForSlides()
    {
        using var document = PresentationDocument.Create(new MemoryStream(), PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation(new SlideIdList());
        var slidePart = presentationPart.AddNewPart<SlidePart>("rId1");
        slidePart.Slide = new P.Slide(new P.CommonSlideData(), new P.Timing());

        RawXmlHelper.Execute(slidePart.Slide, "/p:sld", "append", "<p:transition/>").Should().Be(1);

        var children = XDocument.Parse(slidePart.Slide.OuterXml).Root!.Elements().Select(e => e.Name.LocalName).ToList();
        children.Should().ContainInOrder("cSld", "transition", "timing");
    }

    [Fact]
    public void RawXmlHelper_IsZipUriPath_RecognizesPackageXmlAndRelationshipParts()
    {
        RawXmlHelper.IsZipUriPath("xl/workbook.xml").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("xl/worksheets/sheet1.xml?download=1#frag").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("xl/worksheets/_rels/sheet1.xml.rels").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("/Sheet1").Should().BeFalse();
        RawXmlHelper.IsZipUriPath("xl/media/image1.png").Should().BeFalse();
    }

    [Fact]
    public void WorksheetBloatFilter_FilterSheetStream_RemovesOnlyBareEmptyCells()
    {
        const string xml =
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<dimension ref=\"A1:XFD1048576\"/>" +
            "<sheetData>" +
            "<row r=\"1\" spans=\"1:1\"><c r=\"A1\"/></row>" +
            "<row r=\"2\" spans=\"2:2\"><c r=\"B2\"><v>7</v></c></row>" +
            "<row r=\"3\" spans=\"3:3\"><c r=\"C3\" s=\"1\"/></row>" +
            "<row r=\"4\" spans=\"4:4\"><c r=\"D4\"><f>SUM(B2:B3)</f></c></row>" +
            "<row r=\"1048576\" spans=\"16384:16384\"><c r=\"XFD1048576\"/></row>" +
            "</sheetData>" +
            "</worksheet>";

        using var input = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));
        using var output = new MemoryStream();

        WorksheetBloatFilter.FilterSheetStream(input, output).Should().Be(2);

        output.Position = 0;
        var filtered = XDocument.Load(output);
        var rows = filtered.Descendants(SpreadsheetNs + "row").ToList();
        rows.Select(r => (string?)r.Attribute("r")).Should().Equal("2", "3", "4");
        filtered.Descendants(SpreadsheetNs + "c").Select(c => (string?)c.Attribute("r")).Should().Equal("B2", "C3", "D4");
        filtered.Root!.Element(SpreadsheetNs + "dimension")!.Attribute("ref")!.Value.Should().Be("B2:D4");
    }
}
