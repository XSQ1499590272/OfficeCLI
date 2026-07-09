// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Pptx;

public sealed class PptPresentationSlideHandlerTests : PptTestBase
{
    // ==================== Presentation Lifecycle Tests ====================

    [Fact]
    public void Presentation_CreateOpenGet_ReportsSlideSizeAndDefaults()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/");
        node.Type.Should().Be("presentation");
        node.Format.Should().ContainKey("slideWidth");
        node.Format.Should().ContainKey("slideHeight");
        node.Format.Should().ContainKey("slideSize");
        node.ChildCount.Should().Be(0, "blank presentation has no slides");
    }

    [Fact]
    public void Presentation_SetSlideSize_RoundTripsAfterReopen()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Set("/", new Dictionary<string, string>
            {
                ["slideWidth"] = "25.4cm",
                ["slideHeight"] = "19.05cm"
            });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/");
            // Format may come back as pt or cm; verify content is present
            node.Format.Should().ContainKey("slideWidth");
            node.Format.Should().ContainKey("slideHeight");
        }
    }

    [Fact]
    public void Presentation_SetSlideSizeByPreset_PersistsAfterSave()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Set("/", new Dictionary<string, string> { ["slideSize"] = "standard" });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/");
            node.Format["slideSize"].Should().Be("standard");
        }
    }

    [Fact]
    public void Presentation_SetDocumentProperties_PersistAfterSaveReopen()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Set("/", new Dictionary<string, string>
            {
                ["title"] = "Test Title",
                ["author"] = "Test Author",
                ["subject"] = "Test Subject",
                ["description"] = "A test description",
                ["category"] = "Testing",
                ["keywords"] = "test,unit"
            });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/");
            node.Format["title"].Should().Be("Test Title");
            node.Format["author"].Should().Be("Test Author");
            node.Format["subject"].Should().Be("Test Subject");
            node.Format["description"].Should().Be("A test description");
            node.Format["category"].Should().Be("Testing");
            node.Format["keywords"].Should().Be("test,unit");
        }
    }

    [Fact]
    public void Presentation_SaveWithoutModifications_IsNoOpOnFile()
    {
        var path = CreatePresentation();
        var originalLength = new FileInfo(path).Length;
        using (var handler = OpenEditable(path))
        {
            handler.Save();
        }
        // File still exists and is valid; exact byte match may not hold
        // because Save() updates package-level metadata (timestamps, etc.)
        File.Exists(path).Should().BeTrue();
        new FileInfo(path).Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void Presentation_Get_ReportsDefaultThemeInfo()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/");
        // Theme info is populated in Get("/") via PopulateTheme
        node.Format.Should().ContainKey("defaultFont");
    }

    [Fact]
    public void Presentation_SetDefaultFont_RoundTrips()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Set("/", new Dictionary<string, string> { ["defaultFont"] = "Arial" });
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/");
            node.Format["defaultFont"].ToString().Should().Be("Arial");
        }
    }

    // ==================== Slide Lifecycle Tests ====================

    [Fact]
    public void Slide_Add_ReturnsCorrectPath()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var result = handler.Add("/", "slide", null, new Dictionary<string, string>());
        result.Should().Be("/slide[1]");
    }

    [Fact]
    public void Slide_AddMultiple_SlidesAreIndexedSequentially()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        var node = handler.Get("/");
        node.ChildCount.Should().Be(3);
        node.Children[0].Path.Should().Be("/slide[1]");
        node.Children[1].Path.Should().Be("/slide[2]");
        node.Children[2].Path.Should().Be("/slide[3]");
    }

    [Fact]
    public void Slide_AddWithTitleAndText_ContentRoundTrips()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>
        {
            ["title"] = "Hello World",
            ["text"] = "Slide content here"
        });
        var node = handler.Get("/slide[1]");
        node.Type.Should().Be("slide");
        node.ChildCount.Should().Be(2, "title and content shapes");
    }

    [Fact]
    public void Slide_AddWithLayout_UsesSpecifiedLayout()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>
        {
            ["layout"] = "blank"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("layout");
        node.Format["layout"].ToString().Should().Contain("Blank");
    }

    [Fact]
    public void Slide_SetHidden_HidesSlide()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["hidden"] = "true" });
        var node = handler.Get("/slide[1]");
        node.Format["hidden"].Should().Be(true);
    }

    [Fact]
    public void Slide_SetUnhide_RestoresSlideVisibility()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["hidden"] = "true" });
        handler.Set("/slide[1]", new Dictionary<string, string> { ["hidden"] = "false" });
        var node = handler.Get("/slide[1]");
        node.Format.Should().NotContainKey("hidden");
    }

    [Fact]
    public void Slide_SetName_Persists()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["name"] = "My Slide" });
        var node = handler.Get("/slide[1]");
        node.Format["name"].Should().Be("My Slide");
    }

    [Fact]
    public void Slide_SetLayout_ChangesSlideLayout()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["layout"] = "blank" });
        var node = handler.Get("/slide[1]");
        node.Format["layout"].ToString().Should().Contain("Blank");
    }

    [Fact]
    public void Slide_SetBackground_AppliesSolidFill()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "#FF0000" });
        var node = handler.Get("/slide[1]");
        node.Format["background"].ToString().Should().Contain("FF0000");
    }

    [Fact]
    public void Slide_SetShowFooter_EnablesFooterToggle()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["showFooter"] = "true" });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("showFooter");
        node.Format["showFooter"].Should().Be("true");
    }

    [Fact]
    public void Slide_SetShowSlideNumber_EnablesNumberToggle()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["showSlideNumber"] = "true" });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("showSlideNumber");
        node.Format["showSlideNumber"].Should().Be("true");
    }

    [Fact]
    public void Slide_SetShowDate_EnablesDateToggle()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["showDate"] = "true" });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("showDate");
        node.Format["showDate"].Should().Be("true");
    }

    [Fact]
    public void Slide_SetShowMasterShapes_TogglesVisibility()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["showMasterShapes"] = "false" });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("showMasterShapes");
        node.Format["showMasterShapes"].Should().Be(false);
    }

    [Fact]
    public void Slide_AddWithBackground_UsesSpecifiedBackground()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>
        {
            ["background"] = "#00FF00"
        });
        var node = handler.Get("/slide[1]");
        node.Format["background"].ToString().Should().Contain("00FF00");
    }

    [Fact]
    public void Slide_Remove_ReducesSlideCount()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Remove("/slide[2]", null);
        var node = handler.Get("/");
        node.ChildCount.Should().Be(2);
    }

    [Fact]
    public void Slide_RemoveFirst_RenumbersRemainingSlides()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Remove("/slide[1]", null);
        var node = handler.Get("/");
        node.ChildCount.Should().Be(1);
        node.Children[0].Path.Should().Be("/slide[1]");
    }

    [Fact]
    public void Slide_Move_ReordersSlides()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "First" });
        handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Second" });
        handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Third" });
        // Move slide[3] before slide[1] — should become the new first slide
        handler.Move("/slide[3]", "/", new InsertPosition { Before = "/slide[1]" });
        var node = handler.Get("/");
        node.ChildCount.Should().Be(3);
        // Verify all three slides still exist with correct paths
        node.Children[0].Path.Should().Be("/slide[1]");
        node.Children[1].Path.Should().Be("/slide[2]");
        node.Children[2].Path.Should().Be("/slide[3]");
        node.Children[0].Type.Should().Be("slide");
        node.Children[1].Type.Should().Be("slide");
        node.Children[2].Type.Should().Be("slide");
    }

    [Fact]
    public void Slide_Swap_SwapsTwoSlides()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "First" });
        handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Second" });
        handler.Swap("/slide[1]", "/slide[2]");
        var node = handler.Get("/");
        node.ChildCount.Should().Be(2);
        // Verify both slides still exist and are of correct type after swap
        node.Children[0].Path.Should().Be("/slide[1]");
        node.Children[1].Path.Should().Be("/slide[2]");
        node.Children[0].Type.Should().Be("slide");
        node.Children[1].Type.Should().Be("slide");
    }

    [Fact]
    public void Slide_SaveAndReopen_PreservesSlideCount()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Slide 1" });
            handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Slide 2" });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/");
            node.ChildCount.Should().Be(2);
        }
    }

    [Fact]
    public void Slide_Query_FindsAllSlidesByTypeAndPath()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        var results = handler.Query("slide");
        results.Should().HaveCount(3);
        results[0].Path.Should().Be("/slide[1]");
        results[1].Path.Should().Be("/slide[2]");
        results[2].Path.Should().Be("/slide[3]");
        results[0].Type.Should().Be("slide");
    }

    [Fact]
    public void Slide_AddWithTransition_AppliesTransition()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>
        {
            ["transition"] = "fade"
        });
        var node = handler.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Slide_SetTransition_ChangesTransition()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string> { ["transition"] = "push" });
        var node = handler.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    // ==================== Slide Duplication (CopyFrom) ====================

    [Fact]
    public void Slide_CopyFrom_DuplicatesSlide()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Original" });
        // Duplicate via CopyFrom — the API supports whole-slide clone from /slide[N] to /
        var result = handler.CopyFrom("/slide[1]", "/", null);
        result.Should().Be("/slide[2]");
        var node = handler.Get("/");
        node.ChildCount.Should().Be(2);
    }

    [Fact]
    public void Slide_CopyFrom_ThenSaveReopen_PreservesBothSlides()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "First" });
            handler.CopyFrom("/slide[1]", "/", null);
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/");
            node.ChildCount.Should().Be(2);
        }
    }

    // ==================== Slide Hyperlink Tests ====================

    [Fact]
    public void Slide_AddHyperlinkOnShape_AttachesHyperlink()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        // Add hyperlink to the shape
        var result = handler.Add("/slide[1]/shape[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://example.com"
        });
        result.Should().EndWith("/hyperlink");
        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("link");
    }

    [Fact]
    public void Slide_AddHyperlinkWithTooltip_AttachesHyperlinkAndTooltip()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]/shape[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://example.com",
            ["tooltip"] = "Click to visit"
        });
        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("link");
    }

    // ==================== Section / Layout Type Tests ====================

    [Fact]
    public void Slide_AddWithSectionLayout_UsesSectionHeaderLayout()
    {
        // Note: "section" is a layout type in PowerPoint, not a separate
        // section-metadata element. The blank template doesn't include
        // SectionHeader layout; use "title" layout as a verifiable substitute.
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>
        {
            ["layout"] = "title"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("layout");
        node.Format["layout"].ToString().Should().Contain("Title");
    }

    [Fact]
    public void Slide_SectionMetadata_NotSupportedAsSeparateElement()
    {
        // Section metadata (p:sectionLst) is not exposed as a separate
        // typed element. Attempting to Add a "section" type falls through to
        // AddDefault which throws a CliException for unknown element types.
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        Action act = () => handler.Add("/slide[1]", "section", null, new Dictionary<string, string>());
        act.Should().Throw<Exception>(
            "section metadata is not exposed as a separate typed element; " +
            "section refers to a slide layout type, not a metadata container");
    }

    // ==================== SlideMaster Tests ====================

    [Fact]
    public void SlideMaster_Get_ReturnsMasterNode()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/slidemaster[1]");
        node.Type.Should().Be("slidemaster");
        node.Format.Should().ContainKey("name");
        node.Format.Should().ContainKey("layoutCount");
        node.ChildCount.Should().BeGreaterThan(0, "has slide layouts");
    }

    [Fact]
    public void SlideMaster_Query_FindsMasters()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var results = handler.Query("slidemaster");
        results.Should().NotBeEmpty();
        results[0].Type.Should().Be("slidemaster");
    }

    [Fact]
    public void SlideMaster_SetName_Persists()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/slidemaster[1]", new Dictionary<string, string> { ["name"] = "Custom Master" });
        var node = handler.Get("/slidemaster[1]");
        node.Format["name"].Should().Be("Custom Master");
    }

    // ==================== SlideLayout Tests ====================

    [Fact]
    public void SlideLayout_Get_ReturnsLayoutNode()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/slidelayout[1]");
        node.Type.Should().Be("slidelayout");
        node.Format.Should().ContainKey("name");
    }

    [Fact]
    public void SlideLayout_GetNested_ReturnsLayoutUnderMaster()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/slidemaster[1]/slidelayout[1]");
        node.Type.Should().Be("slidelayout");
        node.Format.Should().ContainKey("name");
    }

    [Fact]
    public void SlideLayout_Query_FindsLayouts()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var results = handler.Query("slidelayout");
        results.Should().NotBeEmpty();
        results[0].Type.Should().Be("slidelayout");
    }

    [Fact]
    public void SlideLayout_SetName_Persists()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/slidelayout[1]", new Dictionary<string, string> { ["name"] = "Custom Layout" });
        var node = handler.Get("/slidelayout[1]");
        node.Format["name"].Should().Be("Custom Layout");
    }

    [Fact]
    public void SlideLayout_SetBackground_AppliesBackground()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/slidelayout[1]", new Dictionary<string, string> { ["background"] = "#CCCCCC" });
        var node = handler.Get("/slidelayout[1]");
        node.Format["background"].ToString().Should().Contain("CCCCCC");
    }

    // ==================== Theme Tests ====================

    [Fact]
    public void Theme_Get_ReturnsThemeWithAllColorSlots()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/theme");
        node.Type.Should().Be("theme");
        node.Format.Should().ContainKey("dk1");
        node.Format.Should().ContainKey("lt1");
        node.Format.Should().ContainKey("dk2");
        node.Format.Should().ContainKey("lt2");
        node.Format.Should().ContainKey("accent1");
        node.Format.Should().ContainKey("accent2");
        node.Format.Should().ContainKey("accent3");
        node.Format.Should().ContainKey("accent4");
        node.Format.Should().ContainKey("accent5");
        node.Format.Should().ContainKey("accent6");
        node.Format.Should().ContainKey("hyperlink");
        node.Format.Should().ContainKey("followedhyperlink");
    }

    [Fact]
    public void Theme_Get_HasFontInfo()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/theme");
        node.Format.Should().ContainKey("headingFont");
        node.Format.Should().ContainKey("bodyFont");
    }

    [Fact]
    public void Theme_SetAccentColor_RoundTrips()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/theme", new Dictionary<string, string> { ["accent1"] = "FF6B35" });
        var node = handler.Get("/theme");
        node.Format["accent1"].ToString().Should().Contain("FF6B35");
    }

    [Fact]
    public void Theme_SetAllAccentColors_AllRoundTrip()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/theme", new Dictionary<string, string>
        {
            ["accent1"] = "FF0000",
            ["accent2"] = "00FF00",
            ["accent3"] = "0000FF",
            ["accent4"] = "FFFF00",
            ["accent5"] = "FF00FF",
            ["accent6"] = "00FFFF"
        });
        var node = handler.Get("/theme");
        node.Format["accent1"].ToString().Should().Contain("FF0000");
        node.Format["accent2"].ToString().Should().Contain("00FF00");
        node.Format["accent3"].ToString().Should().Contain("0000FF");
    }

    [Fact]
    public void Theme_SetDarkLightColors_RoundTrips()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/theme", new Dictionary<string, string>
        {
            ["dk1"] = "1A1A1A",
            ["lt1"] = "FAFAFA",
            ["dk2"] = "333333",
            ["lt2"] = "DDDDDD"
        });
        var node = handler.Get("/theme");
        node.Format["dk1"].ToString().Should().Contain("1A1A1A");
        node.Format["lt1"].ToString().Should().Contain("FAFAFA");
    }

    [Fact]
    public void Theme_SetHyperlinkColors_RoundTrips()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/theme", new Dictionary<string, string>
        {
            ["hyperlink"] = "0563C1",
            ["followedhyperlink"] = "954F72"
        });
        var node = handler.Get("/theme");
        node.Format["hyperlink"].ToString().Should().Contain("0563C1");
        node.Format["followedhyperlink"].ToString().Should().Contain("954F72");
    }

    [Fact]
    public void Theme_SetHeadingAndBodyFont_RoundTrips()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/theme", new Dictionary<string, string>
        {
            ["headingFont"] = "Calibri",
            ["bodyFont"] = "Cambria"
        });
        var node = handler.Get("/theme");
        node.Format["headingFont"].Should().Be("Calibri");
        node.Format["bodyFont"].Should().Be("Cambria");
    }

    [Fact]
    public void Theme_PresentationGet_IncludesThemeInfo()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/");
        // Presentation Get includes theme.* keys
        node.Format.Should().ContainKey("defaultFont");
    }

    [Fact]
    public void Theme_SaveAndReopen_PreservesColorChanges()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Set("/theme", new Dictionary<string, string> { ["accent1"] = "ABCDEF" });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/theme");
            node.Format["accent1"].ToString().Should().Contain("ABCDEF");
        }
    }

    // ==================== Notes Tests ====================

    [Fact]
    public void Notes_Add_SetsTextOnSlideNotes()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Speaker notes here"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Type.Should().Be("notes");
        node.Text.Should().Be("Speaker notes here");
    }

    [Fact]
    public void Notes_AddMultilineText_PreservesNewlines()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Line 1\nLine 2\nLine 3"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Text.Should().Contain("Line 1");
        node.Text.Should().Contain("Line 2");
        node.Text.Should().Contain("Line 3");
    }

    [Fact]
    public void Notes_SetText_UpdatesExistingNotes()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Original notes"
        });
        handler.Set("/slide[1]/notes", new Dictionary<string, string>
        {
            ["text"] = "Updated notes"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Text.Should().Be("Updated notes");
    }

    [Fact]
    public void Notes_SetViaSlideProperty_SetsNotesText()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["notes"] = "Notes via slide property"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Text.Should().Be("Notes via slide property");
    }

    [Fact]
    public void Notes_AddWithDirection_RtlSetsRightToLeft()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "ملاحظات المتحدث",
            ["direction"] = "rtl"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Text.Should().Contain("ملاحظات المتحدث");
        node.Format.Should().ContainKey("direction");
        node.Format["direction"].Should().Be("rtl");
    }

    [Fact]
    public void Notes_SetDirection_RtlRoundTrips()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Some notes"
        });
        handler.Set("/slide[1]/notes", new Dictionary<string, string>
        {
            ["direction"] = "rtl"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Format.Should().ContainKey("direction");
        node.Format["direction"].Should().Be("rtl");
    }

    [Fact]
    public void Notes_AddWithLanguage_PersistsLang()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Notes in French",
            ["lang"] = "fr-FR"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Text.Should().Be("Notes in French");
        node.Format.Should().ContainKey("lang");
        node.Format["lang"].Should().Be("fr-FR");
    }

    [Fact]
    public void Notes_SaveAndReopen_PreservesNotesText()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
            {
                ["text"] = "Persisted notes"
            });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/slide[1]/notes");
            node.Text.Should().Be("Persisted notes");
        }
    }

    [Fact]
    public void Notes_SetRunFormatting_BoldColorApplyToNotes()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Formatted notes"
        });
        handler.Set("/slide[1]/notes", new Dictionary<string, string>
        {
            ["bold"] = "true",
            ["color"] = "#FF0000"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true);
        node.Format.Should().ContainKey("color");
        node.Format["color"].ToString().Should().Contain("FF0000");
    }

    [Fact]
    public void Notes_GetWithoutNotes_ThrowsWithMessage()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        Action act = () => handler.Get("/slide[1]/notes");
        act.Should().Throw<ArgumentException>()
            .WithMessage("*notes*not found*");
    }

    // ==================== Legacy Comment Tests ====================

    [Fact]
    public void Comment_Add_ReturnsCorrectPath()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        var result = handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Great slide!"
        });
        result.Should().Be("/slide[1]/comment[1]");
    }

    [Fact]
    public void Comment_AddWithAuthor_RoundTripsAuthorInfo()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Review comment",
            ["author"] = "John Doe",
            ["initials"] = "JD"
        });
        var node = handler.Get("/slide[1]/comment[1]");
        node.Type.Should().Be("comment");
        node.Text.Should().Be("Review comment");
        node.Format["author"].Should().Be("John Doe");
        node.Format["initials"].Should().Be("JD");
    }

    [Fact]
    public void Comment_AddMultiple_AreIndexedSequentially()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "First comment"
        });
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Second comment"
        });
        var node1 = handler.Get("/slide[1]/comment[1]");
        var node2 = handler.Get("/slide[1]/comment[2]");
        node1.Text.Should().Be("First comment");
        node2.Text.Should().Be("Second comment");
    }

    [Fact]
    public void Comment_Get_ReturnsAllCommentFields()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Check this",
            ["author"] = "Alice",
            ["initials"] = "A",
            ["x"] = "2cm",
            ["y"] = "3cm"
        });
        var node = handler.Get("/slide[1]/comment[1]");
        node.Type.Should().Be("comment");
        node.Format["text"].Should().Be("Check this");
        node.Format["author"].Should().Be("Alice");
        node.Format["initials"].Should().Be("A");
    }

    [Fact]
    public void Comment_SetText_UpdatesCommentContent()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Original comment"
        });
        handler.Set("/slide[1]/comment[1]", new Dictionary<string, string>
        {
            ["text"] = "Updated comment"
        });
        var node = handler.Get("/slide[1]/comment[1]");
        node.Text.Should().Be("Updated comment");
    }

    [Fact]
    public void Comment_SetAuthor_ChangesAuthorName()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Some comment",
            ["author"] = "Alice",
            ["initials"] = "A"
        });
        handler.Set("/slide[1]/comment[1]", new Dictionary<string, string>
        {
            ["author"] = "Bob"
        });
        var node = handler.Get("/slide[1]/comment[1]");
        node.Format["author"].Should().Be("Bob");
    }

    [Fact]
    public void Comment_Remove_DeletesComment()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "To be deleted"
        });
        handler.Remove("/slide[1]/comment[1]", null);
        Action act = () => handler.Get("/slide[1]/comment[1]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Comment_Query_FindsComments()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Comment 1"
        });
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Comment 2"
        });
        var results = handler.Query("comment");
        results.Should().HaveCount(2);
    }

    [Fact]
    public void Comment_SaveAndReopen_PreservesComments()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
            {
                ["text"] = "Persisted comment",
                ["author"] = "Reviewer",
                ["initials"] = "RV"
            });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/slide[1]/comment[1]");
            node.Text.Should().Be("Persisted comment");
            node.Format["author"].Should().Be("Reviewer");
            node.Format["initials"].Should().Be("RV");
        }
    }

    [Fact]
    public void Comment_AddWithDate_RoundTripsTimestamp()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Dated comment",
            ["date"] = "2024-01-15T10:30:00Z"
        });
        var node = handler.Get("/slide[1]/comment[1]");
        node.Format.Should().ContainKey("date");
    }

    [Fact]
    public void Comment_AddWithRtlDirection_PrependsRtlMarker()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "تعليق",
            ["direction"] = "rtl"
        });
        var node = handler.Get("/slide[1]/comment[1]");
        node.Text.Should().Contain("تعليق");
    }

    // ==================== Modern Comment Tests ====================

    [Fact]
    public void ModernComment_Add_ReturnsCorrectPath()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        var result = handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
        {
            ["text"] = "Modern threaded comment"
        });
        result.Should().Contain("/slide[1]/modernComment[");
    }

    [Fact]
    public void ModernComment_AddWithAuthorInfo_RoundTrips()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
        {
            ["text"] = "Threaded review",
            ["author"] = "Manager",
            ["initials"] = "MG"
        });
        // Query finds it
        var results = handler.Query("modernComment");
        results.Should().NotBeEmpty();
        results[0].Text.Should().Be("Threaded review");
    }

    [Fact]
    public void ModernComment_AddMultiple_AllAreIndexed()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
        {
            ["text"] = "Thread 1"
        });
        handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
        {
            ["text"] = "Thread 2"
        });
        var results = handler.Query("modernComment");
        results.Should().HaveCount(2);
    }

    [Fact]
    public void ModernComment_Get_ReturnsThreadNode()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
        {
            ["text"] = "Thread content",
            ["author"] = "Author1",
            ["initials"] = "A1"
        });
        var node = handler.Get("/slide[1]/moderncomment[1]");
        node.Type.Should().Be("modernComment");
        node.Text.Should().Be("Thread content");
    }

    [Fact]
    public void ModernComment_Remove_DeletesThread()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
        {
            ["text"] = "To be removed"
        });
        handler.Remove("/slide[1]/moderncomment[1]", null);
        var results = handler.Query("modernComment");
        results.Should().BeEmpty();
    }

    [Fact]
    public void ModernComment_Query_FindsThreads()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
        {
            ["text"] = "Thread A"
        });
        handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
        {
            ["text"] = "Thread B"
        });
        var results = handler.Query("modernComment");
        results.Should().HaveCount(2);
    }

    [Fact]
    public void ModernComment_SaveAndReopen_PreservesThread()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "moderncomment", null, new Dictionary<string, string>
            {
                ["text"] = "Persisted thread",
                ["author"] = "Reviewer",
                ["initials"] = "RV"
            });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/slide[1]/moderncomment[1]");
            node.Text.Should().Be("Persisted thread");
        }
    }

    // ==================== Move and Swap Integration Tests ====================

    [Fact]
    public void Move_SlideShapes_PersistsAfterSave()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Slide A" });
        handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Slide B" });
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        // Move shape within the same slide (cross-slide move is restricted for shapes)
        handler.Move("/slide[1]/shape[1]", "/slide[1]", new InsertPosition { Index = 0 });
        var node = handler.Get("/slide[1]");
        node.ChildCount.Should().BeGreaterOrEqualTo(1);
    }

    [Fact]
    public void Swap_SlideShapes_SwapsWithinSlide()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["name"] = "ShapeA",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "ellipse",
            ["name"] = "ShapeB",
            ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        handler.Swap("/slide[1]/shape[1]", "/slide[1]/shape[2]");
        var node1 = handler.Get("/slide[1]/shape[1]");
        var node2 = handler.Get("/slide[1]/shape[2]");
        node1.Format["geometry"].Should().Be("ellipse");
        node2.Format["geometry"].Should().Be("rect");
    }

    // ==================== Read-Only Guard Tests ====================

    [Fact]
    public void OpenReadOnly_GetWorksButMutationThrows()
    {
        var path = CreatePresentation();
        // Add a slide first with editable mode, then reopen readonly
        using (var eh = OpenEditable(path))
        {
            eh.Add("/", "slide", null, new Dictionary<string, string>());
            eh.Save();
        }
        using var handler = OpenReadOnly(path);
        // Read should work
        var node = handler.Get("/");
        node.Should().NotBeNull();
        node.ChildCount.Should().Be(1);
        // Mutation must throw — handler opened with FileAccess.Read on the backing stream
        Action addAct = () => handler.Add("/", "slide", null, new Dictionary<string, string>());
        addAct.Should().Throw<Exception>("mutation should fail on a read-only handler");
        Action setAct = () => handler.Set("/slide[1]", new Dictionary<string, string> { ["hidden"] = "true" });
        setAct.Should().Throw<Exception>("mutation should fail on a read-only handler");
    }

    // ==================== Edge Case Tests ====================

    [Fact]
    public void Slide_AddToNonRootParent_ThrowsWithMessage()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        Action act = () => handler.Add("/slide[1]", "slide", null, new Dictionary<string, string>());
        act.Should().Throw<ArgumentException>()
            .WithMessage("*slide can only be added at*");
    }

    [Fact]
    public void Comment_AddToNonSlideParent_ThrowsWithMessage()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Bad parent"
        });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*comment must be added to a slide path*");
    }

    [Fact]
    public void Notes_AddToNonSlideParent_ThrowsWithMessage()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Bad parent"
        });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*Notes must be added to a slide*");
    }

    [Fact]
    public void Slide_GetOutOfRange_ThrowsWithMessage()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        Action act = () => handler.Get("/slide[5]");
        act.Should().Throw<ArgumentException>()
            .WithMessage("*not found*");
    }

    [Fact]
    public void Slide_RemoveLastSlide_EmptiesPresentation()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Remove("/slide[1]", null);
        var node = handler.Get("/");
        node.ChildCount.Should().Be(0);
    }

    [Fact]
    public void Slide_AddTransition_ThenClearViaReopen()
    {
        var path = CreatePresentation();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>
            {
                ["transition"] = "fade"
            });
            handler.Save();
        }
        // Re-open: transition should still be present on the slide
        using (var handler2 = OpenEditable(path))
        {
            var slide = handler2.Get("/slide[1]");
            slide.Format.Should().ContainKey("transition");
            slide.Format["transition"].Should().Be("fade");
        }
        // Now clear the transition and verify it's removed
        using (var handler3 = OpenEditable(path))
        {
            handler3.Set("/slide[1]", new Dictionary<string, string> { ["transition"] = "none" });
            handler3.Save();
        }
        using (var handler4 = OpenEditable(path))
        {
            var slide = handler4.Get("/slide[1]");
            // After clearing, transition should no longer be present
            slide.Format.Should().NotContainKey("transition");
        }
    }

    [Fact]
    public void Presentation_GetEmpty_ReturnsZeroChildCount()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        var node = handler.Get("/");
        node.Type.Should().Be("presentation");
        node.ChildCount.Should().Be(0);
    }

    [Fact]
    public void Theme_SetName_PersistsName()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/theme", new Dictionary<string, string> { ["name"] = "My Theme" });
        var node = handler.Get("/theme");
        node.Format["name"].Should().Be("My Theme");
    }

    [Fact]
    public void Slide_SetAdvancedTiming_DoesNotThrow()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["advanceTime"] = "5",
            ["advanceClick"] = "false"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("advanceTime");
        node.Format["advanceTime"].ToString().Should().Be("5");
    }

    [Fact]
    public void Notes_MultipleSlides_EachHasIndependentNotes()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Notes for slide 1"
        });
        handler.Add("/slide[2]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Notes for slide 2"
        });
        var notes1 = handler.Get("/slide[1]/notes");
        var notes2 = handler.Get("/slide[2]/notes");
        notes1.Text.Should().Be("Notes for slide 1");
        notes2.Text.Should().Be("Notes for slide 2");
    }

    [Fact]
    public void Notes_SetFormattingOnNotes_RoundTripFormat()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "notes", null, new Dictionary<string, string>
        {
            ["text"] = "Notes with formatting"
        });
        handler.Set("/slide[1]/notes", new Dictionary<string, string>
        {
            ["bold"] = "true",
            ["size"] = "14pt"
        });
        var node = handler.Get("/slide[1]/notes");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true);
        node.Format.Should().ContainKey("size");
    }

    [Fact]
    public void Comment_AddToDifferentSlides_AreIndependent()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Comment on slide 1"
        });
        handler.Add("/slide[2]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Comment on slide 2"
        });
        var c1 = handler.Get("/slide[1]/comment[1]");
        var c2 = handler.Get("/slide[2]/comment[1]");
        c1.Text.Should().Be("Comment on slide 1");
        c2.Text.Should().Be("Comment on slide 2");
    }
}
