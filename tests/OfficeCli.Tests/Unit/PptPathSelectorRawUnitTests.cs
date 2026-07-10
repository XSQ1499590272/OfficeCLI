// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using P = DocumentFormat.OpenXml.Presentation;

using OfficeCli.Tests.Pptx;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class PptPathSelectorRawUnitTests : PptTestBase
{
    private string CreatePresentationWithSlide()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
        }
        return path;
    }

    // ==================== Selector Tests ====================

    [Fact]
    public void AttributeFilter_ParseAndApply_CoversSelectorOperators()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/slide[1]/shape[1]",
                Type = "shape",
                Text = "Revenue Title",
                Format = new()
                {
                    ["text"] = "Revenue",
                    ["x"] = "1cm",
                    ["y"] = "2cm",
                    ["width"] = "5cm",
                    ["height"] = "3cm"
                }
            },
            new()
            {
                Path = "/slide[1]/shape[2]",
                Type = "shape",
                Text = "Cost Subtitle",
                Format = new()
                {
                    ["text"] = "Cost",
                    ["x"] = "3cm",
                    ["width"] = "4cm"
                }
            }
        };

        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[text=Revenue]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[1]");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[text!=Revenue]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[2]");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[text~=Rev]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[1]");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[text~=r\"Cost$\"]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[2]");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[width>\"4cm\"]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[1]");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[y]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[1]");
    }

    [Fact]
    public void AttributeFilter_TypeSelectors_MatchShapeTypes()
    {
        var nodes = new List<DocumentNode>
        {
            new() { Path = "/slide[1]/shape[1]", Type = "shape", Format = new() { ["kind"] = "shape" } },
            new() { Path = "/slide[1]/picture[1]", Type = "picture", Format = new() { ["kind"] = "picture" } },
            new() { Path = "/slide[1]/table[1]", Type = "table", Format = new() { ["kind"] = "table" } },
            new() { Path = "/slide[1]/chart[1]", Type = "chart", Format = new() { ["kind"] = "chart" } },
            new() { Path = "/slide[1]/group[1]", Type = "group", Format = new() { ["kind"] = "group" } }
        };

        // Use expression tree for type-based filtering via ParseExpr + ApplyExpr
        AttributeFilter.ApplyExpr(nodes, AttributeFilter.ParseExpr("shape[kind=shape]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[1]");
        AttributeFilter.ApplyExpr(nodes, AttributeFilter.ParseExpr("picture[kind=picture]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/picture[1]");
        AttributeFilter.ApplyExpr(nodes, AttributeFilter.ParseExpr("table[kind=table]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/table[1]");
    }

    [Fact]
    public void AttributeFilter_ExpressionTree_SupportsAndOrForShapeSelectors()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/slide[1]/shape[1]",
                Type = "shape",
                Text = "Bold Revenue",
                Format = new() { ["bold"] = true, ["text"] = "Revenue" }
            },
            new()
            {
                Path = "/slide[1]/shape[2]",
                Type = "shape",
                Text = "Cost Subtitle",
                Format = new() { ["text"] = "Cost" }
            },
            new()
            {
                Path = "/slide[1]/shape[3]",
                Type = "shape",
                Text = "Bold Summary",
                Format = new() { ["bold"] = true, ["text"] = "Summary" }
            }
        };

        // AND: bold AND revenue text
        AttributeFilter.ApplyExpr(nodes,
            AttributeFilter.ParseExpr("shape[bold=true and text=Revenue]"))
            .Select(n => n.Path)
            .Should().Equal("/slide[1]/shape[1]");

        // OR: bold OR cost text
        AttributeFilter.ApplyExpr(nodes,
            AttributeFilter.ParseExpr("shape[bold=true or text=Cost]"))
            .Select(n => n.Path)
            .Should().Equal("/slide[1]/shape[1]", "/slide[1]/shape[2]", "/slide[1]/shape[3]");
    }

    [Fact]
    public void AttributeFilter_ReportsUnknownKeysAndNonNumericComparisons()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/slide[1]/shape[1]",
                Type = "shape",
                Text = "Revenue",
                Format = new()
                {
                    ["bold"] = true,
                    ["width"] = "5cm"
                }
            }
        };

        var expr = AttributeFilter.ParseExpr("shape[blod=true and width>wide]");
        var (results, warnings) = AttributeFilter.ApplyExprWithWarnings(nodes, expr);

        results.Should().BeEmpty();
        warnings.Should().ContainSingle(w => w.Kind == "unknown_key" && w.Key == "blod");
        warnings.Should().ContainSingle(w => w.Kind == "non_numeric" && w.Value == "wide");
        warnings.Single(w => w.Kind == "unknown_key").Available.Should().Contain("bold");
    }

    [Fact]
    public void AttributeFilter_NestedGroupTraversal_SupportsGroupChildSelectors()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/slide[1]/group[1]/shape[1]",
                Type = "shape",
                Text = "Inner A",
                Format = new() { ["text"] = "Alpha" }
            },
            new()
            {
                Path = "/slide[1]/group[1]/shape[2]",
                Type = "shape",
                Text = "Inner B",
                Format = new() { ["text"] = "Beta" }
            },
            new()
            {
                Path = "/slide[1]/shape[1]",
                Type = "shape",
                Text = "Top Level",
                Format = new() { ["text"] = "Gamma" }
            }
        };

        // Select shapes with text "Alpha" regardless of nesting
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[text=Alpha]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/group[1]/shape[1]");
    }

    [Fact]
    public void AttributeFilter_FullPathAtNameConnectorTargets_ResolveCorrectly()
    {
        var nodes = new List<DocumentNode>
        {
            new()
            {
                Path = "/slide[1]/connector[1]",
                Type = "connector",
                Format = new()
                {
                    ["startShape"] = "/slide[1]/shape[@name=BoxA]",
                    ["endShape"] = "/slide[1]/shape[@name=BoxB]"
                }
            },
            new()
            {
                Path = "/slide[1]/shape[1]",
                Type = "shape",
                Format = new() { ["name"] = "BoxA" }
            },
            new()
            {
                Path = "/slide[1]/shape[2]",
                Type = "shape",
                Format = new() { ["name"] = "BoxB" }
            }
        };

        // Select shapes by @name= attribute
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[name=BoxA]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[1]");
        AttributeFilter.Apply(nodes, AttributeFilter.Parse("shape[name=BoxB]"))
            .Select(n => n.Path).Should().Equal("/slide[1]/shape[2]");
    }

    // ==================== Path Tests ====================

    [Fact]
    public void Get_SlidePath_ResolvesWithCorrectType()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        var node = handler.Get("/slide[1]");
        node.Path.Should().Be("/slide[1]");
        node.Type.Should().Be("slide");
    }

    [Fact]
    public void Get_RootPath_ResolvesPresentation()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/");
        node.Path.Should().Be("/");
        node.Type.Should().Be("presentation");
        node.Format.Should().ContainKey("slideWidth");
        node.Format.Should().ContainKey("slideHeight");
    }

    [Fact]
    public void Get_SlideMasterPath_ResolvesWithCorrectType()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/slideMaster[1]");
        node.Path.Should().Be("/slidemaster[1]");
        node.Type.Should().Be("slidemaster");
    }

    [Fact]
    public void Get_SlideLayoutPath_ResolvesWithCorrectType()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/slideLayout[1]");
        node.Path.Should().Be("/slidelayout[1]");
        node.Type.Should().Be("slidelayout");
    }

    [Fact]
    public void Get_SlideNotesPath_ThrowsWhenNoNotes()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Blank presentations have no notes, so Get("/slide[1]/notes") throws.
        Action act = () => handler.Get("/slide[1]/notes");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Get_SlideNotesPath_ResolvesAfterAddingNotes()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string> { ["text"] = "Speaker note" });
        var node = handler.Get("/slide[1]/notes");
        node.Type.Should().Be("notes");
    }

    [Fact]
    public void Get_ThemePath_ResolvesTheme()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/theme");
        node.Path.Should().Be("/theme");
        node.Type.Should().Be("theme");
    }

    [Fact]
    public void Get_SlideMasterPath_ContainsLayoutChildren()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/slideMaster[1]", depth: 2);
        node.Path.Should().Be("/slidemaster[1]");
        node.Children.Should().NotBeNull();
        node.Children!.Should().Contain(c => c.Type == "slidelayout");
    }

    [Fact]
    public void Get_SlideShapePath_ResolvesWithCorrectType()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "Hello" });
        var node = handler.Get("/slide[1]/shape[1]");
        node.Type.Should().Be("textbox");
        node.Text.Should().Be("Hello");
    }

    [Fact]
    public void Get_SlideGroupShapePath_ResolvesNestedShape()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Create an empty group with geometry
        handler.Add("/slide[1]", "group", null, new Dictionary<string, string>
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "5cm"
        });
        handler.Add("/slide[1]/group[1]", "shape", null, new Dictionary<string, string> { ["text"] = "Nested" });
        var node = handler.Get("/slide[1]/group[1]/shape[1]");
        node.Type.Should().Be("textbox");
        node.Text.Should().Be("Nested");
    }

    [Fact]
    public void Get_SlideTableRowCellPath_ResolvesCell()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "table", null, new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "2" });
        // Tables use tr/tc path segment names and @id-based indexing in output
        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Type.Should().Be("tc");
    }

    // ==================== Raw XML Tests: raw action ====================

    [Fact]
    public void Raw_PresentationPart_ReturnsNonEmptyXml()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var xml = handler.Raw("/presentation");
        xml.Should().NotBeNullOrEmpty();
        xml.Should().Contain("<p:presentation");
    }

    [Fact]
    public void Raw_SlidePart_ReturnsNonEmptyXml()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        var xml = handler.Raw("/slide[1]");
        xml.Should().NotBeNullOrEmpty();
        xml.Should().Contain("<p:sld");
    }

    [Fact]
    public void Raw_SlideMasterPart_ReturnsNonEmptyXml()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var xml = handler.Raw("/slideMaster[1]");
        xml.Should().NotBeNullOrEmpty();
        xml.Should().Contain("<p:sldMaster");
    }

    [Fact]
    public void Raw_SlideLayoutPart_ReturnsNonEmptyXml()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var xml = handler.Raw("/slideLayout[1]");
        xml.Should().NotBeNullOrEmpty();
        xml.Should().Contain("<p:sldLayout");
    }

    [Fact]
    public void Raw_ThemePart_ReturnsNonEmptyXml()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var xml = handler.Raw("/theme");
        xml.Should().NotBeNullOrEmpty();
        xml.Should().Contain("<a:theme");
    }

    [Fact]
    public void Raw_NoteSlidePart_ThrowsWhenNoNotes()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Blank pptx with a slide still has no notes -- expecting ArgumentException
        Action act = () => handler.Raw("/noteSlide[1]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Raw_NotesMasterPart_ThrowsWhenNoNotesMaster()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Raw("/notesMaster");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Raw_NoteSlidePart_ReturnsNonEmptyXmlAfterAddingNotes()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string> { ["text"] = "Speaker note" });
        var xml = handler.Raw("/noteSlide[1]");
        xml.Should().NotBeNullOrEmpty();
        xml.Should().Contain("<p:notes");
    }

    [Fact]
    public void Raw_ChartPart_ReturnsNonEmptyXmlAfterAddingChart()
    {
        var path = CreatePresentationWithSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
            {
                ["data"] = "Series1:1,2,3"
            });
            handler.Save();
        }
        // The chart part is stored in the zip; find it by listing OPC entries
        // that contain "chart" in their path, then read via Raw.
        using var archive = System.IO.Compression.ZipFile.OpenRead(path);
        var chartEntry = archive.Entries
            .FirstOrDefault(e => e.FullName.Contains("chart", StringComparison.OrdinalIgnoreCase)
                                 && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        chartEntry.Should().NotBeNull("should have a chart XML part in the zip");
        var chartPath = "/" + chartEntry!.FullName;
        using var handler2 = OpenReadOnly(path);
        var xml = handler2.Raw(chartPath);
        xml.Should().NotBeNullOrEmpty();
        xml.Should().Contain("chartSpace");
    }

    // ==================== Raw XML Tests: raw-set actions ====================

    [Fact]
    public void RawSet_Append_AddsElementToSlide()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);

        handler.RawSet("/slide[1]", "/p:sld/p:cSld/p:spTree", "append",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"999\" name=\"Appended\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Appended Text</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        var xml = handler.Raw("/slide[1]");
        xml.Should().Contain("Appended");
    }

    [Fact]
    public void RawSet_Prepend_AddsElementAtBeginningOfSlide()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);

        handler.RawSet("/slide[1]", "/p:sld/p:cSld/p:spTree", "prepend",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"888\" name=\"Prepended\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Prepended Text</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        var xml = handler.Raw("/slide[1]");
        xml.Should().Contain("Prepended");
    }

    [Fact]
    public void RawSet_Replace_ReplacesElementInSlide()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);

        // First append a shape, then replace it
        handler.RawSet("/slide[1]", "/p:sld/p:cSld/p:spTree", "append",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"777\" name=\"Original\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original Text</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        handler.RawSet("/slide[1]", "//p:sp[p:nvSpPr/p:cNvPr[@name='Original']]", "replace",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"777\" name=\"Original\" customAttr=\"keep\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"100000\" y=\"100000\"/><a:ext cx=\"200000\" cy=\"200000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Replaced Text</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        var xml = handler.Raw("/slide[1]");
        xml.Should().Contain("Replaced");
        xml.Should().Contain("customAttr=\"keep\"");
    }

    [Fact]
    public void RawSet_Remove_RemovesElementFromSlide()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);

        handler.RawSet("/slide[1]", "/p:sld/p:cSld/p:spTree", "append",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"666\" name=\"ToRemove\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Remove Me</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        handler.RawSet("/slide[1]", "//p:sp[p:nvSpPr/p:cNvPr[@name='ToRemove']]", "remove", null);

        var xml = handler.Raw("/slide[1]");
        xml.Should().NotContain("Remove Me");
    }

    [Fact]
    public void RawSet_SetAttr_SetsAttributeOnElement()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);

        handler.RawSet("/slide[1]", "/p:sld/p:cSld/p:spTree", "append",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"555\" name=\"AttrTarget\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Attr Target Text</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        handler.RawSet("/slide[1]", "//p:sp[p:nvSpPr/p:cNvPr[@name='AttrTarget']]/p:nvSpPr/p:cNvPr", "setattr",
            "descr=\"Added attribute\"");

        var xml = handler.Raw("/slide[1]");
        xml.Should().Contain("descr=");
        xml.Should().Contain("Added attribute");
    }

    [Fact]
    public void RawSet_InsertBefore_InsertsElementBeforeTarget()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);

        // Append a target shape
        handler.RawSet("/slide[1]", "/p:sld/p:cSld/p:spTree", "append",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"444\" name=\"Target\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Target Shape</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        // Insert before the target
        handler.RawSet("/slide[1]", "//p:sp[p:nvSpPr/p:cNvPr[@name='Target']]", "insertbefore",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"443\" name=\"BeforeTarget\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Before Insertion</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        var xml = handler.Raw("/slide[1]");
        xml.Should().Contain("Before Insertion");
        xml.Should().Contain("Target Shape");
    }

    [Fact]
    public void RawSet_InsertAfter_InsertsElementAfterTarget()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);

        // Append a target shape
        handler.RawSet("/slide[1]", "/p:sld/p:cSld/p:spTree", "append",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"333\" name=\"Anchor\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Anchor Shape</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        // Insert after the anchor
        handler.RawSet("/slide[1]", "//p:sp[p:nvSpPr/p:cNvPr[@name='Anchor']]", "insertafter",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<p:nvSpPr><p:cNvPr id=\"334\" name=\"AfterAnchor\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100000\" cy=\"100000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>After Insertion</a:t></a:r></a:p></p:txBody>" +
            "</p:sp>");

        var xml = handler.Raw("/slide[1]");
        xml.Should().Contain("After Insertion");
        xml.Should().Contain("Anchor Shape");
    }

    [Fact]
    public void RawSet_OnPresentationPart_AppendsToPresentation()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);

        handler.RawSet("/presentation", "/p:presentation", "append",
            "<p:custDataLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
            "<p:custData id=\"testData\"/>" +
            "</p:custDataLst>");

        var xml = handler.Raw("/presentation");
        xml.Should().Contain("custDataLst");
    }

    [Fact]
    public void RawSet_OnSlideMasterPart_AppendsToSlideMaster()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);

        handler.RawSet("/slideMaster[1]", "/p:sldMaster", "append",
            "<p:custDataLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
            "<p:custData id=\"masterData\"/>" +
            "</p:custDataLst>");

        var xml = handler.Raw("/slideMaster[1]");
        xml.Should().Contain("custDataLst");
    }

    [Fact]
    public void RawSet_OnSlideLayoutPart_AppendsToSlideLayout()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);

        handler.RawSet("/slideLayout[1]", "/p:sldLayout", "append",
            "<p:custDataLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
            "<p:custData id=\"layoutData\"/>" +
            "</p:custDataLst>");

        var xml = handler.Raw("/slideLayout[1]");
        xml.Should().Contain("custDataLst");
    }

    [Fact]
    public void RawSet_OnThemePart_AppendsToTheme()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);

        handler.RawSet("/theme", "/a:theme", "append",
            "<a:custExtLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<a:custExt uri=\"{test}\"/>" +
            "</a:custExtLst>");

        var xml = handler.Raw("/theme");
        xml.Should().Contain("custExtLst");
    }

    [Fact]
    public void RawXmlHelper_Execute_CoversCoreMutationActionsWithNamespacePreservation()
    {
        var slide = new P.Slide(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.Shape(new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 2, Name = "Shape2" },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(),
                        new P.TextBody()),
                    new P.Shape(new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 4, Name = "Shape4" },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(),
                        new P.TextBody()))));

        RawXmlHelper.Execute(slide, "/p:sld/p:cSld/p:spTree", "prepend",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:nvSpPr><p:cNvPr id=\"1\" name=\"S1\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody/></p:sp>")
            .Should().Be(1);
        RawXmlHelper.Execute(slide, "/p:sld/p:cSld/p:spTree", "append",
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:nvSpPr><p:cNvPr id=\"5\" name=\"S5\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody/></p:sp>")
            .Should().Be(1);

        var shapes = slide.CommonSlideData!.ShapeTree!.Elements<P.Shape>().ToList();
        shapes.Select(s => s.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id!.Value)
            .Should().Equal(1u, 2u, 4u, 5u);

        var xdoc = XDocument.Parse(slide.OuterXml);
        var pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        xdoc.Root!.Name.NamespaceName.Should().Be(pNs);
    }

    // ==================== Help / Schema Tests ====================

    [Fact]
    public void SchemaHelp_AllPptxSchemaElements_RenderWithJson()
    {
        var schemaDir = Path.Combine(RepoRoot(), "schemas", "help", "pptx");
        var jsonFiles = Directory.GetFiles(schemaDir, "*.json");
        jsonFiles.Should().NotBeEmpty("schemas/help/pptx/ should contain schema JSON files");

        foreach (var file in jsonFiles)
        {
            var elementName = Path.GetFileNameWithoutExtension(file);
            var result = RunCliOk("help", "pptx", elementName, "--json");
            result.Stdout.Should().NotBeNullOrEmpty($"help pptx {elementName} --json should produce output");
            result.Stdout.Should().Contain("\"element\"",
                $"help pptx {elementName} --json should contain \"element\"");
        }
    }

    // ==================== Negative Tests ====================

    [Fact]
    public void Get_InvalidSlideIndex_ThrowsArgumentException()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Get("/slide[99]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Get_InvalidSlideIndexZero_ThrowsArgumentException()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Get("/slide[0]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Get_InvalidSlideMasterIndex_ThrowsArgumentException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Get("/slideMaster[99]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Get_InvalidSlideLayoutIndex_ThrowsArgumentException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Get("/slideLayout[99]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Raw_MissingPart_ThrowsArgumentException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Raw("/nonexistent");
        act.Should().Throw<ArgumentException>()
            .WithMessage("*Unknown part*");
    }

    [Fact]
    public void Raw_MissingZipUriPart_ThrowsArgumentException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Raw("/ppt/slides/nonexistent.xml");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void RawSet_InvalidPartPath_ThrowsArgumentException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.RawSet("/nonexistent", "/root", "append", "<x/>");
        act.Should().Throw<ArgumentException>()
            .WithMessage("*Unknown part*");
    }

    [Fact]
    public void RawSet_InvalidAction_ThrowsOnValidation()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.RawSet("/slide[1]", "/p:sld", "invalid_action", "<x/>");
        act.Should().Throw<ArgumentException>()
            .WithMessage("*Unknown action*");
    }

    [Fact]
    public void Get_UnsupportedElementTypeName_ThrowsArgumentException()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Get("/slide[1]/unknown[1]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Query_UnknownElementType_ReturnsEmpty()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var results = handler.Query("nonexistentelement");
        results.Should().BeEmpty();
    }

    [Fact]
    public void Query_PathStyleSelector_ThrowsArgumentException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Query("/slide");
        act.Should().Throw<ArgumentException>()
            .WithMessage("*path-style selectors starting with '/' are not allowed*");
    }

    [Fact]
    public void Raw_NullPartPath_ThrowsArgumentNullException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Raw(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RawSet_NullPartPath_ThrowsArgumentNullException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.RawSet(null!, "/root", "append", "<x/>");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RawSet_NullXPath_ThrowsArgumentNullException()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.RawSet("/slide[1]", null!, "append", "<x/>");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RawSet_NullAction_ThrowsArgumentNullException()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.RawSet("/slide[1]", "/root", null!, "<x/>");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RawXmlHelper_IsZipUriPath_RecognizesPptxParts()
    {
        RawXmlHelper.IsZipUriPath("ppt/slides/slide1.xml").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("ppt/slideMasters/slideMaster1.xml").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("ppt/slideLayouts/slideLayout1.xml").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("ppt/theme/theme1.xml").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("ppt/presentation.xml").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("ppt/notesSlides/notesSlide1.xml").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("ppt/slides/_rels/slide1.xml.rels").Should().BeTrue();
        RawXmlHelper.IsZipUriPath("/slide[1]").Should().BeFalse();
        RawXmlHelper.IsZipUriPath("ppt/media/image1.png").Should().BeFalse();
    }

    [Fact]
    public void RawXmlHelper_Execute_ReordersKnownSchemaChildrenForSlides()
    {
        using var stream = new MemoryStream();
        using var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new P.Presentation(new P.SlideIdList());
        var slidePart = presentationPart.AddNewPart<SlidePart>("rId1");
        slidePart.Slide = new P.Slide(new P.CommonSlideData(), new P.Timing());

        RawXmlHelper.Execute(slidePart.Slide, "/p:sld", "append", "<p:transition/>").Should().Be(1);

        var children = XDocument.Parse(slidePart.Slide.OuterXml).Root!.Elements()
            .Select(e => e.Name.LocalName).ToList();
        children.Should().ContainInOrder("cSld", "transition", "timing");
    }
}
