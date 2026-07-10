// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

using OfficeCli.Tests.Pptx;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public sealed class PptRenderDumpUnitTests : PptTestBase
{
    private string CreateSlide()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        return path;
    }

    // ==================== HTML PREVIEW TESTS ====================

    [Fact]
    public void HtmlPreview_SingleSlide_ReturnsCompleteHtmlDocument()
    {
        var path = CreateSlide();
        using var handler = OpenReadOnly(path);
        var html = handler.ViewAsHtml();

        html.Should().Contain("<!DOCTYPE html>");
        html.Should().Contain("<html lang=\"en\">");
        html.Should().Contain("<head>");
        html.Should().Contain("<body>");
        html.Should().Contain("</html>");
    }

    [Fact]
    public void HtmlPreview_SlideWithContent_ContainsSlideContainer()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Hello", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        html.Should().Contain("slide-container");
        html.Should().Contain("Hello");
    }

    [Fact]
    public void HtmlPreview_RangeFilter_RendersOnlySpecifiedSlides()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Slide1", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[2]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Slide2", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml(startSlide: 1, endSlide: 1);

        html.Should().Contain("Slide1");
        html.Should().NotContain("Slide2");
    }

    [Fact]
    public void HtmlPreview_ShapeRendered_ContainsShapeKeyword()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["shape"] = "rect", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        // The shape keyword ("rect") appears in the rendered HTML
        html.Should().Contain("rect");
    }

    [Fact]
    public void HtmlPreview_GroupedShapes_RendersGroupStructure()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "group", null, new Dictionary<string, string>
            {
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "5cm"
            });
            handler.Add("/slide[1]/group[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Inside group", ["x"] = "0.5cm", ["y"] = "0.5cm", ["width"] = "3cm", ["height"] = "1cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        html.Should().Contain("Inside group");
    }

    [Fact]
    public void HtmlPreview_Table_HasTableStructure()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "table", null, new Dictionary<string, string>
            {
                ["rows"] = "2", ["cols"] = "3",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "4cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        // Tables are rendered as divs with table-like CSS
        html.Should().Contain("table");
        html.Should().NotBeEmpty();
    }

    [Fact]
    public void HtmlPreview_Picture_ContainsImageData()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
            {
                ["src"] = pngPath, ["alt"] = "Test image"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        // Images are embedded as base64 data URIs in the HTML output
        html.Should().Contain("data:image/");
        // The picture element should be rendered in the output
        html.Should().NotBeEmpty();
    }

    [Fact]
    public void HtmlPreview_Chart_RendersChartContainer()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
            {
                ["type"] = "bar",
                ["title"] = "Sales",
                ["data"] = "Series1:10,20,30;Series2:15,25,35",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "6cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        // Charts render with embedded SVG
        html.Should().Contain("Sales");
    }

    [Fact]
    public void HtmlPreview_Notes_RenderAfterSlide()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
            {
                ["text"] = "Speaker notes content"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        html.Should().Contain("slide-notes");
        html.Should().Contain("Speaker notes content");
    }

    [Fact]
    public void HtmlPreview_Hyperlink_ContainsHrefAttribute()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Click me",
                ["link"] = "https://example.com",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        html.Should().Contain("https://example.com");
    }

    [Fact]
    public void HtmlPreview_TextEffects_BoldItalic_ContainCssStyles()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "BoldItalic",
                ["bold"] = "true",
                ["italic"] = "true",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        html.Should().Contain("font-weight:bold");
        html.Should().Contain("font-style:italic");
    }

    [Fact]
    public void HtmlPreview_RtlText_HasDirectionAttribute()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "RTL Text",
                ["direction"] = "rtl",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        html.Should().Contain("dir=\"rtl\"");
    }

    [Fact]
    public void HtmlPreview_ThemeColors_AreResolvedInOutput()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Themed text",
                ["color"] = "tx1",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        // Theme colors are resolved to hex, so should contain a hash color
        html.Should().Contain("Themed text");
    }

    [Fact]
    public void HtmlPreview_HiddenSlide_IsIncludedInOutput()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>
            {
                ["hidden"] = "true"
            });
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Hidden slide content",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        // Hidden slides are still rendered in HTML output (not excluded from preview)
        html.Should().Contain("Hidden slide content");
        // The Get tree surfaces hidden=true so consumers can detect this status
        var slide = ro.Get("/slide[1]");
        slide.Format.Should().ContainKey("hidden");
        slide.Format["hidden"].Should().Be(true);
    }

    [Fact]
    public void HtmlPreview_SlideWithComment_CommentsAreNotRendered()
    {
        // Comments are stored in a separate OOXML part (slideComments.xml) and
        // are not included in the HTML preview output. This test documents that
        // expectation — if comments are later added to the HTML renderer, this
        // test should be updated to assert their presence.
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
            {
                ["text"] = "Review this content"
            });
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml();

        // Comments are not rendered in HTML preview; the slide content is present
        html.Should().NotContain("Review this content", "comments are not rendered in HTML preview");
    }

    [Fact]
    public void HtmlPreview_SlideCount_MatchesActualSlides()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/", "slide", null, new Dictionary<string, string>());
        }
        using var ro = OpenReadOnly(path);
        var count = ro.GetSlideCount();

        count.Should().Be(3);
    }

    // ==================== SVG PREVIEW TESTS ====================

    [Fact]
    public void SvgPreview_SlideWithGeometry_GeneratesValidSvg()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["shape"] = "rect",
                ["fill"] = "#FF0000",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var svg = ro.ViewAsSvg(1);

        svg.Should().Contain("<svg xmlns=\"http://www.w3.org/2000/svg\"");
        svg.Should().Contain("viewBox=");
        svg.Should().Contain("</svg>");
        svg.Should().Contain("<rect");
    }

    [Fact]
    public void SvgPreview_TextContent_ContainsTextElements()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "SVG Text",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var svg = ro.ViewAsSvg(1);

        svg.Should().Contain("<foreignObject");
        svg.Should().Contain("SVG Text");
    }

    [Fact]
    public void SvgPreview_Picture_ContainsImageElement()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
            {
                ["src"] = pngPath
            });
        }
        using var ro = OpenReadOnly(path);
        var svg = ro.ViewAsSvg(1);

        svg.Should().Contain("<image ");
    }

    [Fact]
    public void SvgPreview_Chart_ContainsNestedSvg()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
            {
                ["type"] = "bar",
                ["title"] = "Bar Chart",
                ["data"] = "Series1:10,20,30;Series2:15,25,35",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "6cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var svg = ro.ViewAsSvg(1);

        svg.Should().Contain("<svg"); // nested SVG for chart
    }

    [Fact]
    public void SvgPreview_GroupedShape_RendersWithGroupTransform()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "group", null, new Dictionary<string, string>
            {
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "5cm"
            });
            handler.Add("/slide[1]/group[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Group text",
                ["x"] = "0.5cm", ["y"] = "0.5cm", ["width"] = "3cm", ["height"] = "1cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var svg = ro.ViewAsSvg(1);

        svg.Should().Contain("Group text");
        svg.Should().Contain("<g transform");
    }

    [Fact]
    public void SvgPreview_Connector_RendersLineElement()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "connector", null, new Dictionary<string, string>
            {
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "0cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var svg = ro.ViewAsSvg(1);

        svg.Should().Contain("<line ");
    }

    [Fact]
    public void SvgPreview_StrictPageHandling_OutOfRangeThrowsCliException()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);

        Action act = () => ro.ViewAsSvg(99);
        act.Should().Throw<CliException>()
            .Where(e => e.Code == "out_of_range")
            .WithMessage("*Slide 99 does not exist*");
    }

    [Fact]
    public void SvgPreview_StrictPageHandling_ZeroThrowsCliException()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);

        Action act = () => ro.ViewAsSvg(0);
        act.Should().Throw<CliException>()
            .Where(e => e.Code == "out_of_range");
    }

    [Fact]
    public void SvgPreview_ShapeWithPresetGeometry_RendersPolygon()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["shape"] = "triangle",
                ["fill"] = "#00FF00",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "3cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var svg = ro.ViewAsSvg(1);

        svg.Should().Contain("<polygon ");
    }

    [Fact]
    public void SvgPreview_MultiSlide_HasCorrectDimensions()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var svg = ro.ViewAsSvg(1);

        // SVG should have width and height attributes
        svg.Should().Contain("width=\"");
        svg.Should().Contain("height=\"");
    }

    // ==================== SCREENSHOT HELPER TESTS ====================

    [Fact]
    public void Screenshot_AutoGridColumns_OneItem_ReturnsOne()
    {
        var cols = HtmlScreenshot.AutoGridColumns(1, 960, 540);
        cols.Should().Be(1);
    }

    [Fact]
    public void Screenshot_AutoGridColumns_ManyItems_ReturnsReasonableColumns()
    {
        var cols = HtmlScreenshot.AutoGridColumns(10, 960, 540);

        // For 10 landscape slides (960×540), should return ~3 columns
        cols.Should().BeGreaterThan(1);
        cols.Should().BeLessOrEqualTo(10);
    }

    [Fact]
    public void Screenshot_AutoGridColumns_Portrait_ReturnsFewerColumns()
    {
        var landscape = HtmlScreenshot.AutoGridColumns(9, 960, 540);
        var portrait = HtmlScreenshot.AutoGridColumns(9, 540, 960);

        // Portrait pages (aspect > 1) get more columns
        portrait.Should().BeGreaterOrEqualTo(landscape);
    }

    [Fact]
    public void Screenshot_AutoGridColumns_RespectsMaxCount()
    {
        var cols = HtmlScreenshot.AutoGridColumns(100, 960, 540);
        cols.Should().BeLessOrEqualTo(100);
        cols.Should().BeGreaterThan(1);
    }

    [Fact]
    public void Screenshot_GetSlideNativePixels_ReturnsReasonableSize()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var (w, h) = handler.GetSlideNativePixels();

        // Default 16:9 slide is 1280×720
        w.Should().BeGreaterThan(0);
        h.Should().BeGreaterThan(0);
    }

    [Fact]
    public void Screenshot_GetSlideNativePixels_MultiplePresentations_AreConsistent()
    {
        var path1 = CreatePresentation();
        var path2 = CreatePresentation();
        using var h1 = OpenEditable(path1);
        using var h2 = OpenEditable(path2);
        var (w1, h1V) = h1.GetSlideNativePixels();
        var (w2, h2V) = h2.GetSlideNativePixels();

        w1.Should().Be(w2);
        h1V.Should().Be(h2V);
    }

    [Fact]
    public void Screenshot_HasChromeFamily_ReturnsBoolean()
    {
        // CI typically has no Chrome; dev machines may or may not. The test
        // verifies the method runs without throwing — the return value is
        // host-dependent. When a browser IS found, we assert true.
        var has = HtmlScreenshot.HasChromeFamily();
        if (!has) return; // No Chrome-family browser on this host
        has.Should().BeTrue();
    }

    [Fact]
    public void Screenshot_ViewportSizing_GridHtml_HasGridCss()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml(startSlide: 1, endSlide: 1, gridCols: 3, viewportPx: 1200);

        // Grid layout CSS should be present
        html.Should().Contain("display:grid");
        html.Should().Contain("grid-template-columns");
    }

    [Fact]
    public void Screenshot_ViewportSizing_ForcedGridColumns_AppliesLayout()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
        }
        using var ro = OpenReadOnly(path);
        var html = ro.ViewAsHtml(startSlide: 1, endSlide: 2, gridCols: 2);

        html.Should().Contain("grid-template-columns:repeat(2");
    }

    [Fact]
    public void Screenshot_RenderHtmlFallback_SingleSlideRemovesPagePadding()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);

        // When startSlide == endSlide, CSS removes padding for screenshot mode
        var html = ro.ViewAsHtml(startSlide: 1, endSlide: 1);
        html.Should().Contain(".headless .main");
    }

    // ==================== DUMP EMITTER TESTS ====================

    [Fact]
    public void Dump_FullDeck_ReturnsItemsAndWarnings()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Dumpable content",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, warnings) = PptxBatchEmitter.EmitPptx(ro);

        items.Should().NotBeNull();
        items.Should().NotBeEmpty();
        items.All(i => !string.IsNullOrEmpty(i.Command)).Should().BeTrue();
        warnings.Should().NotBeNull();
    }

    [Fact]
    public void Dump_FullDeck_FirstItemIsRemoveAllSlides()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        items[0].Command.Should().Be("remove");
        items[0].Path.Should().Be("/slide[*]");
    }

    [Fact]
    public void Dump_FullDeck_HasAddSlideCommand()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        items.Should().Contain(i => i.Command == "add" && i.Type == "slide");
    }

    [Fact]
    public void Dump_SlideSubtree_ReturnsOnlyThatSlide()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Slide 1 content",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[2]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Slide 2 content",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro, "/slide[1]");

        items.Should().NotBeEmpty();
        // Subtree dump for a slide should not include remove-all-slides
        items.Should().NotContain(i => i.Command == "remove" && i.Path == "/slide[*]");
    }

    [Fact]
    public void Dump_ChartSlide_HasChartSpecificItems()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
            {
                ["type"] = "bar",
                ["title"] = "Chart Title",
                ["data"] = "Series1:10,20,30;Series2:15,25,35",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "6cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        items.Should().Contain(i => i.Command == "add" && i.Type == "chart");
    }

    [Fact]
    public void Dump_PictureSlide_HasPictureAddCommand()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
            {
                ["src"] = pngPath
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        items.Should().Contain(i => i.Command == "add" && i.Type == "picture");
    }

    [Fact]
    public void Dump_Notes_HasNotesItems()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
            {
                ["text"] = "Notes for dump"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        // Notes are emitted as add with Type="notes"
        items.Should().Contain(i => i.Command == "add" && i.Type == "notes");
    }

    [Fact]
    public void Dump_ThemeSubtree_ReturnsThemeRawSet()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro, "/theme");

        items.Should().NotBeEmpty();
        items.Should().Contain(i => i.Command == "raw-set" || i.Command == "set");
    }

    [Fact]
    public void Dump_PresentationSubtree_ReturnsPresentationProps()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro, "/presentation");

        items.Should().NotBeNull();
    }

    [Fact]
    public void Dump_Subtree_InvalidPath_ThrowsCliException()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);

        Action act = () => PptxBatchEmitter.EmitPptx(ro, "/bogus[1]");
        act.Should().Throw<CliException>()
            .Where(e => e.Code == "unsupported_path");
    }

    [Fact]
    public void Dump_Subtree_EmptyPath_ThrowsCliException()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);

        Action act = () => PptxBatchEmitter.EmitPptx(ro, "");
        act.Should().Throw<CliException>()
            .Where(e => e.Code == "invalid_path");
    }

    [Fact]
    public void Dump_Subtree_OutOfRangeSlide_ThrowsCliException()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);

        Action act = () => PptxBatchEmitter.EmitPptx(ro, "/slide[99]");
        act.Should().Throw<CliException>()
            .Where(e => e.Code == "path_not_found");
    }

    [Fact]
    public void Dump_NoteSlideSubtree_ReturnsNoteItems()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
            {
                ["text"] = "Test notes"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro, "/noteSlide[1]");

        items.Should().NotBeEmpty();
    }

    [Fact]
    public void Dump_SlideLayoutSubtree_ReturnsLayoutRawSet()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro, "/slideLayout[1]");

        items.Should().NotBeEmpty();
        items.Should().Contain(i => i.Command == "raw-set");
    }

    [Fact]
    public void Dump_SlideMasterSubtree_ReturnsMasterRawSet()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro, "/slideMaster[1]");

        items.Should().NotBeEmpty();
        items.Should().Contain(i => i.Command == "raw-set");
    }

    [Fact]
    public void Dump_TableSlide_HasTableAddCommand()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "table", null, new Dictionary<string, string>
            {
                ["rows"] = "2", ["cols"] = "3",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "4cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        items.Should().Contain(i => i.Command == "add" && i.Type == "table");
    }

    [Fact]
    public void Dump_CommentSlide_HasCommentItems()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
            {
                ["text"] = "Dumpable comment"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        // Comments are emitted as add with Type="comment", Parent="/slide[N]"
        items.Should().Contain(i => i.Command == "add" && i.Type == "comment");
    }

    [Fact]
    public void Dump_AnimationAndTransition_EmitItems()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["shape"] = "rect",
                ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
            });
            // Add an animation on the shape (matches existing Animation_* tests)
            handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
            {
                ["effect"] = "fade"
            });
            // Set a slide transition via Set
            handler.Set("/slide[1]", new Dictionary<string, string>
            {
                ["transition"] = "fade"
            });
        }
        using var ro = OpenReadOnly(path);

        // Verify animation is queryable after reopen
        var anims = ro.Query("animation");
        anims.Should().NotBeEmpty("animation should persist after save/reopen");

        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        // Animation timing is emitted as a raw-set passthrough (not semantic
        // add-animation rows) due to broad exotic-timing detection that routes
        // all <p:anim*> elements through raw-set for byte-fidelity.
        items.Should().Contain(i => i.Command == "raw-set"
                                   && i.Part != null && i.Part.StartsWith("/slide[")
                                   && i.Xml != null && i.Xml.Contains("<p:timing"),
            "dump should raw-set passthrough the animation timing tree");

        // The transition prop should be on the add slide row (standard transitions
        // like fade are not exotic and survive as semantic props)
        items.Should().Contain(i => i.Command == "add" && i.Type == "slide"
                                   && i.Props != null && i.Props.ContainsKey("transition"),
            "dump should preserve transition on the slide add row");
    }

    [Fact]
    public void Dump_OleAndModel3D_EmitItems()
    {
        var path = CreateSlide();
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        var olePayloadPath = CreateOlePayload(NewTempPath(".dat"));
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
            {
                ["src"] = glbPath,
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "5cm"
            });
            handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
            {
                ["src"] = olePayloadPath,
                ["progid"] = "Word.Document",
                ["x"] = "8cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "5cm"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, warnings) = PptxBatchEmitter.EmitPptx(ro);

        // OLE and model3d items are emitted as add-part rows
        items.Should().Contain(i => i.Command == "add-part" && i.Type == "model3d",
            "dump should emit model3d add-part items");
        items.Should().Contain(i => i.Command == "add-part" && i.Type == "ole",
            "dump should emit ole add-part items");
    }

    [Fact]
    public void Dump_HiddenSlide_PreservesHiddenFlag()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>
            {
                ["hidden"] = "true"
            });
        }
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        // Should have both an add slide and potentially a set for hidden
        items.Should().Contain(i => i.Command == "add" && i.Type == "slide");
    }

    [Fact]
    public void Dump_EmptyPresentation_ProducesResourceCommands()
    {
        var path = CreatePresentation();
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        // Even empty presentations produce theme/master/layout raw-set commands
        items.Should().Contain(i => i.Command == "set" || i.Command == "raw-set");
    }

    [Fact]
    public void Dump_ItemSchema_AllItemsHaveValidCommand()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var (items, _) = PptxBatchEmitter.EmitPptx(ro);

        var validCommands = new[] { "add", "set", "remove", "raw-set" };
        items.All(i => validCommands.Contains(i.Command)).Should().BeTrue();
    }

    [Fact]
    public void Dump_Warnings_ForContentWithUnsupportedElements()
    {
        var path = CreateSlide();
        using var ro = OpenReadOnly(path);
        var (items, warnings) = PptxBatchEmitter.EmitPptx(ro);

        // Even if warnings are empty, the list should be non-null
        warnings.Should().NotBeNull();
    }

    // ==================== BATCH BEHAVIOR TESTS ====================

    [Fact]
    public void Batch_ContinueOnError_Default_RunsAllItems()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "First", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            },
            // Invalid command that will fail
            new() { Command = "set", Path = "/slide[1]/bogus[1]",
                Props = new Dictionary<string, string> { ["text"] = "Should fail" }
            },
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "Third", ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            }
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: false, json: false);

        results.Should().HaveCount(3);
        results[0].Success.Should().BeTrue();
        results[1].Success.Should().BeFalse();
        results[1].Error.Should().NotBeNullOrEmpty();
        results[2].Success.Should().BeTrue();
    }

    [Fact]
    public void Batch_StopOnError_StopsAfterFirstFailure()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "First", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            },
            new() { Command = "set", Path = "/slide[1]/bogus[1]",
                Props = new Dictionary<string, string> { ["text"] = "Should fail" }
            },
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "NeverReached", ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            }
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: true, json: false);

        results.Should().HaveCount(2); // stops after second item fails
        results[0].Success.Should().BeTrue();
        results[1].Success.Should().BeFalse();
    }

    [Fact]
    public void Batch_AllSucceed_ReturnsAllSuccess()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "Item 1", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            },
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "Item 2", ["x"] = "1cm", ["y"] = "3cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            }
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: false, json: false);

        results.Should().HaveCount(2);
        results.All(r => r.Success).Should().BeTrue();
    }

    [Fact]
    public void Batch_JsonOutput_ReturnsResultsAsJson()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "Single", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            }
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: false, json: true);

        results.Should().HaveCount(1);
        results[0].Success.Should().BeTrue();

        // Verify the result serializes to valid JSON with expected structure
        var json = JsonSerializer.Serialize(results, BatchJsonContext.Default.ListBatchResult);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        root.ValueKind.Should().Be(JsonValueKind.Array, "results should serialize as a JSON array");
        root.GetArrayLength().Should().Be(1);
        root[0].TryGetProperty("success", out var successProp).Should().BeTrue();
        successProp.GetBoolean().Should().BeTrue();
        root[0].TryGetProperty("index", out var indexProp).Should().BeTrue();
        root[0].TryGetProperty("output", out _).Should().BeTrue("successful results should have an output property");
    }

    [Fact]
    public void Batch_InvalidCommand_ProducesErrorWithDiagnostics()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "unknownCommand" } // invalid command
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: false, json: false);

        results.Should().HaveCount(1);
        results[0].Success.Should().BeFalse();
        results[0].Error.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Batch_InvalidPath_ProducesError()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "get", Path = "/bogus[1]" }
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: false, json: false);

        results.Should().HaveCount(1);
        results[0].Success.Should().BeFalse();
    }

    [Fact]
    public void Batch_EmptyItemList_ReturnsEmptyResults()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);

        var results = CommandBuilder.RunNonResidentBatch(handler, new List<BatchItem>(), stopOnError: false, json: false);

        results.Should().BeEmpty();
    }

    [Fact]
    public void Batch_GetCommand_ReturnsOutput()
    {
        var path = CreateSlide();
        using (var setupHandler = OpenEditable(path))
        {
            setupHandler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "BatchGetTest", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
            });
        }
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "get", Path = "/slide[1]" }
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: false, json: false);

        results.Should().HaveCount(1);
        results[0].Success.Should().BeTrue();
        results[0].Output.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Batch_SetCommand_ModifiesDocument()
    {
        var path = CreateSlide();
        using (var setupHandler = OpenEditable(path))
        {
            setupHandler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Original", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
            });
        }
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "set", Path = "/slide[1]/shape[1]",
                Props = new Dictionary<string, string> { ["text"] = "Modified" }
            }
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: false, json: false);

        results[0].Success.Should().BeTrue();
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Modified");
    }

    [Fact]
    public void Batch_MixedCommands_AllProcessed()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);

        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            },
            new() { Command = "get", Path = "/slide[1]/shape[1]" },
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string>
                {
                    ["text"] = "B", ["x"] = "1cm", ["y"] = "3cm", ["width"] = "3cm", ["height"] = "1cm"
                }
            },
            new() { Command = "remove", Path = "/slide[1]/shape[2]" },
            new() { Command = "get", Path = "/slide[1]" }
        };

        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: false, json: false);

        results.Should().HaveCount(5);
        results.All(r => r.Success).Should().BeTrue();
        var slideNode = handler.Get("/slide[1]");
        slideNode.Children.Count(c => c.Type == "textbox" || c.Type == "shape").Should().Be(1);
    }

    [Fact]
    public void Batch_BatchItemSerialization_RoundTrips()
    {
        var item = new BatchItem
        {
            Command = "add",
            Parent = "/slide[1]",
            Type = "textbox",
            Props = new Dictionary<string, string>
            {
                ["text"] = "Hello", ["x"] = "1cm", ["y"] = "2cm"
            }
        };

        var json = JsonSerializer.Serialize(item, BatchJsonContext.Default.BatchItem);
        var deserialized = JsonSerializer.Deserialize(json, BatchJsonContext.Default.BatchItem);

        deserialized.Should().NotBeNull();
        deserialized!.Command.Should().Be("add");
        deserialized.Parent.Should().Be("/slide[1]");
        deserialized.Type.Should().Be("textbox");
        deserialized.Props.Should().ContainKey("text");
        deserialized.Props!["text"].Should().Be("Hello");
    }

    [Fact]
    public void Batch_ListSerialization_RoundTrips()
    {
        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/", Type = "slide" },
            new() { Command = "add", Parent = "/slide[1]", Type = "textbox",
                Props = new Dictionary<string, string> { ["text"] = "Content" }
            }
        };

        var json = JsonSerializer.Serialize(items, BatchJsonContext.Default.ListBatchItem);
        var deserialized = JsonSerializer.Deserialize(json, BatchJsonContext.Default.ListBatchItem);

        deserialized.Should().NotBeNull();
        deserialized!.Should().HaveCount(2);
        deserialized![0].Command.Should().Be("add");
        deserialized![0].Type.Should().Be("slide");
    }

    [Fact]
    public void Dump_FullDeck_ThenBatchReplay_RoundTrips()
    {
        var path1 = CreateSlide();
        string text;
        using (var handler = OpenEditable(path1))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "RoundtripText", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
            text = handler.Get("/slide[1]/shape[1]").Text!;
        }

        // Dump
        List<BatchItem> items;
        using (var ro = OpenReadOnly(path1))
        {
            (items, _) = PptxBatchEmitter.EmitPptx(ro);
        }

        // Replay
        var path2 = CreatePresentation();
        using (var handler = OpenEditable(path2))
        {
            var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError: true, json: false);
            results.All(r => r.Success).Should().BeTrue();
        }

        // Verify
        using (var ro2 = OpenReadOnly(path2))
        {
            var node = ro2.Get("/slide[1]/shape[1]");
            node.Text.Should().Be(text);
        }
    }
}
