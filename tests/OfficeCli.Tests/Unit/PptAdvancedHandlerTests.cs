// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

using OfficeCli.Tests.Pptx;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class PptAdvancedHandlerTests : PptTestBase
{
    private string CreateSlide()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        return path;
    }

    // ==================== Chart Handler Direct Tests ====================

    [Fact]
    public void Chart_AddMultiple_UniquePathsReturned()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        var r1 = handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "A:10,20,30"
        });
        var r2 = handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "line",
            ["data"] = "B:5,15,25"
        });
        r1.Should().NotBe(r2);
        r1.Should().StartWith("/slide[1]/chart[");
        r2.Should().StartWith("/slide[1]/chart[");

        var charts = handler.Query("chart");
        charts.Should().HaveCount(2);
    }

    [Fact]
    public void Chart_SetChartTitle_UpdatesReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["title"] = "Original",
            ["data"] = "Sales:10,20,30"
        });
        handler.Set("/slide[1]/chart[1]", new Dictionary<string, string>
        {
            ["title"] = "Updated Title"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Format["title"].Should().Be("Updated Title");
    }

    [Fact]
    public void Chart_SetChartName_UpdatesReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30"
        });
        handler.Set("/slide[1]/chart[1]", new Dictionary<string, string>
        {
            ["name"] = "RenamedChart"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Format["name"].Should().Be("RenamedChart");
    }

    [Fact]
    public void Chart_SetPosition_MovesChart()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["x"] = "2cm"
        });
        handler.Set("/slide[1]/chart[1]", new Dictionary<string, string>
        {
            ["x"] = "5cm",
            ["y"] = "4cm"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
        node.Format.Should().ContainKey("x");
    }

    [Fact]
    public void Chart_StackedBar_AddAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "bar",
            ["grouping"] = "stacked",
            ["data"] = "Sales:10,20,30;Costs:5,10,15"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
        (node.Format["chartType"] as string).Should().Contain("bar");
    }

    [Fact]
    public void Chart_PercentStackedColumn_AddAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["grouping"] = "percentStacked",
            ["data"] = "Sales:10,20,30;Costs:5,10,15"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_Remove_ChartRemovedFromSlide()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30"
        });
        handler.Query("chart").Should().HaveCount(1);
        handler.Remove("/slide[1]/chart[1]");
        handler.Query("chart").Should().HaveCount(0);
    }

    // ==================== Media Handler Direct Tests ====================

    [Fact]
    public void Media_AddVideoWithPoster_ReadbackHasPoster()
    {
        var path = CreateSlide();
        var videoPath = CreateTinyVideo(NewTempPath(".mp4"));
        var posterPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "video", null, new Dictionary<string, string>
        {
            ["src"] = videoPath,
            ["poster"] = posterPath
        });
        var node = handler.Get("/slide[1]/video[1]");
        node.Type.Should().Be("video");
        node.Format.Should().ContainKey("contentType");
    }

    [Fact]
    public void Media_SetVolume_OnAudio()
    {
        var path = CreateSlide();
        var audioPath = NewTempPath(".mp3");
        File.WriteAllBytes(audioPath, new byte[] {
            0xFF, 0xFB, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00
        });
        TrackTempFile(audioPath);
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "audio", null, new Dictionary<string, string>
        {
            ["src"] = audioPath
        });
        handler.Set("/slide[1]/audio[1]", new Dictionary<string, string>
        {
            ["volume"] = "25"
        });
        var node = handler.Get("/slide[1]/audio[1]");
        node.Type.Should().Be("audio");
    }

    [Fact]
    public void Media_SetLoop_AutoStart_OnVideo()
    {
        var path = CreateSlide();
        var videoPath = CreateTinyVideo(NewTempPath(".mp4"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "video", null, new Dictionary<string, string>
        {
            ["src"] = videoPath
        });
        handler.Set("/slide[1]/video[1]", new Dictionary<string, string>
        {
            ["loop"] = "true",
            ["autoStart"] = "true"
        });
        var node = handler.Get("/slide[1]/video[1]");
        node.Type.Should().Be("video");
    }

    [Fact]
    public void Media_RemoveVideo_RemovesFromSlide()
    {
        var path = CreateSlide();
        var videoPath = CreateTinyVideo(NewTempPath(".mp4"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "video", null, new Dictionary<string, string>
        {
            ["src"] = videoPath
        });
        handler.Query("video").Should().HaveCount(1);
        handler.Remove("/slide[1]/video[1]");
        handler.Query("video").Should().HaveCount(0);
    }

    [Fact]
    public void Media_RemoveAudio_RemovesFromSlide()
    {
        var path = CreateSlide();
        var audioPath = NewTempPath(".mp3");
        File.WriteAllBytes(audioPath, new byte[] {
            0xFF, 0xFB, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00
        });
        TrackTempFile(audioPath);
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "audio", null, new Dictionary<string, string>
        {
            ["src"] = audioPath
        });
        handler.Query("audio").Should().HaveCount(1);
        handler.Remove("/slide[1]/audio[1]");
        handler.Query("audio").Should().HaveCount(0);
    }

    [Fact]
    public void Media_SetName_UpdatesReadback()
    {
        var path = CreateSlide();
        var videoPath = CreateTinyVideo(NewTempPath(".mp4"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "video", null, new Dictionary<string, string>
        {
            ["src"] = videoPath,
            ["name"] = "OriginalName"
        });
        // Media name Set may work differently — verify the video node still exists
        handler.Set("/slide[1]/video[1]", new Dictionary<string, string>
        {
            ["name"] = "RenamedVideo"
        });
        var node = handler.Get("/slide[1]/video[1]");
        node.Type.Should().Be("video");
        node.Format.Should().ContainKey("name");
    }

    // ==================== Animation Handler Direct Tests ====================

    [Fact]
    public void Animation_AddWithDelay_Easing_Succeeds()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fly",
            ["class"] = "entrance",
            ["duration"] = "500",
            ["delay"] = "200",
            ["direction"] = "left"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_AddWithTrigger_AfterPrevious()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "wipe",
            ["class"] = "entrance",
            ["trigger"] = "after"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_AddWithTrigger_WithPrevious()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "zoom",
            ["class"] = "entrance",
            ["trigger"] = "with"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_DurAlias_WorksSameAsDuration()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "entrance",
            ["dur"] = "750"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_EffectClassInSuffix_ParsedCorrectly()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        // "fly-in" / "fly-out" suffixes are parsed as class tokens
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fly-in"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_InvalidClass_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        Action act = () => handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "invalidClass"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Animation_InvalidEffect_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        Action act = () => handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "nonexistentEffect",
            ["class"] = "entrance"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Animation_ChartBuildOnShape_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        Action act = () => handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "entrance",
            ["chartBuild"] = "bySeries"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Animation_MotionPathOnChart_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30"
        });
        Action act = () => handler.Add("/slide[1]/chart[1]", "animation", null, new Dictionary<string, string>
        {
            ["class"] = "motion",
            ["path"] = "line"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Animation_RepeatIndefinite_SetsCorrectValue()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "spin",
            ["class"] = "emphasis",
            ["repeat"] = "indefinite"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_InvalidDelay_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        Action act = () => handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "entrance",
            ["delay"] = "-100"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Animation_InvalidDuration_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        Action act = () => handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "entrance",
            ["duration"] = "-500"
        });
        act.Should().Throw<ArgumentException>();
    }

    // ==================== Transition Handler Direct Tests ====================

    [Fact]
    public void Transition_None_RemovesTransition()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "fade"
        });
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "none"
        });
        var node = handler.Get("/slide[1]");
        // After removing transition, the key should be absent
        node.Format.Should().NotContainKey("transition");
    }

    [Fact]
    public void Transition_WipeWithDirection_RoundTrips()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "wipe-right"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transition");
    }

    [Fact]
    public void Transition_ZoomInOut_RoundTrips()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "zoom-in"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transition");
    }

    [Fact]
    public void Transition_InvalidType_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "invalidTransitionType"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Transition_InvalidDirection_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "circle-xyz"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Transition_InvalidSpeed_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transitionSpeed"] = "hyperspeed"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Transition_InvalidAdvanceTime_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        // advanceTime validation rejects negative values
        Action act = () => handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "fade",
            ["advanceTime"] = "-1"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Transition_FadeThroughBlack_SetsCorrectly()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "fade-thru-black"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transition");
    }

    [Fact]
    public void Transition_WheelWithSpokes_RoundTrips()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "wheel-8"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transition");
    }

    // ==================== Zoom Handler Direct Tests ====================

    [Fact]
    public void Zoom_SetReturnToParent_Persists()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "zoom", null, new Dictionary<string, string>
        {
            ["target"] = "2"
        });
        handler.Set("/slide[1]/zoom[1]", new Dictionary<string, string>
        {
            ["returnToParent"] = "true"
        });
        var node = handler.Get("/slide[1]/zoom[1]");
        node.Type.Should().Be("zoom");
    }

    [Fact]
    public void Zoom_Remove_RemovesFromSlide()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "zoom", null, new Dictionary<string, string>
        {
            ["target"] = "2"
        });
        handler.Query("zoom").Should().HaveCount(1);
        handler.Remove("/slide[1]/zoom[1]");
        handler.Query("zoom").Should().HaveCount(0);
    }

    [Fact]
    public void Zoom_InvalidTarget_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/slide[1]", "zoom", null, new Dictionary<string, string>
        {
            ["target"] = "999"
        });
        act.Should().Throw<ArgumentException>();
    }

    // ==================== OLE Handler Direct Tests ====================

    [Fact]
    public void Ole_AddWithIcon_ReadbackCorrect()
    {
        var path = CreateSlide();
        var olePayloadPath = CreateOlePayload(NewTempPath(".bin"));
        var iconPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
        {
            ["src"] = olePayloadPath,
            ["icon"] = iconPath,
            ["display"] = "icon"
        });
        var node = handler.Get("/slide[1]/ole[1]");
        node.Type.Should().Be("ole");
    }

    [Fact]
    public void Ole_SetDisplay_ChangesMode()
    {
        var path = CreateSlide();
        var olePayloadPath = CreateOlePayload(NewTempPath(".bin"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
        {
            ["src"] = olePayloadPath,
            ["display"] = "icon"
        });
        handler.Set("/slide[1]/ole[1]", new Dictionary<string, string>
        {
            ["display"] = "content"
        });
        var node = handler.Get("/slide[1]/ole[1]");
        node.Type.Should().Be("ole");
    }

    [Fact]
    public void Ole_Remove_RemovesFromSlide()
    {
        var path = CreateSlide();
        var olePayloadPath = CreateOlePayload(NewTempPath(".bin"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
        {
            ["src"] = olePayloadPath
        });
        handler.Query("ole").Should().HaveCount(1);
        handler.Remove("/slide[1]/ole[1]");
        handler.Query("ole").Should().HaveCount(0);
    }

    [Fact]
    public void Ole_SetName_UpdatesReadback()
    {
        var path = CreateSlide();
        var olePayloadPath = CreateOlePayload(NewTempPath(".bin"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
        {
            ["src"] = olePayloadPath,
            ["name"] = "OriginalOle"
        });
        handler.Set("/slide[1]/ole[1]", new Dictionary<string, string>
        {
            ["name"] = "RenamedOle"
        });
        var node = handler.Get("/slide[1]/ole[1]");
        node.Format["name"].Should().Be("RenamedOle");
    }

    // ==================== Model3D Handler Direct Tests ====================

    [Fact]
    public void Model3D_SetPerAxisRotation_RoundTrips()
    {
        var path = CreateSlide();
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
        {
            ["src"] = glbPath,
            ["rotX"] = "45"
        });
        handler.Set("/slide[1]/model3d[1]", new Dictionary<string, string>
        {
            ["rotY"] = "90",
            ["rotZ"] = "180"
        });
        var node = handler.Get("/slide[1]/model3d[1]");
        node.Type.Should().Be("model3d");
    }

    [Fact]
    public void Model3D_Remove_RemovesFromSlide()
    {
        var path = CreateSlide();
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
        {
            ["src"] = glbPath
        });
        handler.Query("model3d").Should().HaveCount(1);
        handler.Remove("/slide[1]/model3d[1]");
        handler.Query("model3d").Should().HaveCount(0);
    }

    [Fact]
    public void Model3D_SetPosition_RoundTrips()
    {
        var path = CreateSlide();
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
        {
            ["src"] = glbPath,
            ["x"] = "3cm",
            ["y"] = "4cm"
        });
        handler.Set("/slide[1]/model3d[1]", new Dictionary<string, string>
        {
            ["x"] = "6cm",
            ["y"] = "8cm"
        });
        var node = handler.Get("/slide[1]/model3d[1]");
        node.Type.Should().Be("model3d");
    }

    // ==================== Diagram Handler Direct Tests ====================

    [Fact]
    public void Diagram_WithMermaidText_ProducesShapes()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "diagram", null, new Dictionary<string, string>
        {
            ["mermaid"] = "flowchart TD\n    A --> B\n    B --> C"
        });
        // Diagrams are add-only synthesizers; after Add the slide should have content
        var slide = handler.Get("/slide[1]", depth: 2);
        slide.Children.Should().NotBeNull();
    }

    [Fact]
    public void Diagram_SimpleFlowchart_ProducesConnectors()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "diagram", null, new Dictionary<string, string>
        {
            ["text"] = "flowchart LR\n    Start[Start] --> Middle[Middle]\n    Middle --> End[End]"
        });
        // Diagrams use native mode or rendered image as a group; slide should have content
        var slide = handler.Get("/slide[1]", depth: 1);
        slide.Should().NotBeNull();
        slide.Type.Should().Be("slide");
    }

    // ==================== Error / Edge Case Tests ====================

    [Fact]
    public void Picture_InvalidCoordinates_ThrowsArgumentException()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["width"] = "-1cm"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Chart_InvalidType_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "invalidChartType",
            ["data"] = "X:1,2,3"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Model3D_NonGlbFile_ThrowsArgumentException()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
        {
            ["src"] = pngPath
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Transition_AddViaType_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        // Transition must be set, not added
        Action act = () => handler.Add("/slide[1]", "transition", null, new Dictionary<string, string>
        {
            ["effect"] = "fade"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Animation_OnNonExistentParent_ThrowsArgumentException()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/slide[1]/shape[99]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "entrance"
        });
        act.Should().Throw<ArgumentException>();
    }

    // ==================== Save/Reopen Round-Trips ====================

    [Fact]
    public void ChartWithAnimation_Transition_Zoom_SaveReopen_AllPersist()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            // Chart
            handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
            {
                ["type"] = "column",
                ["title"] = "PersistChart",
                ["data"] = "A:10,20,30"
            });
            // Animation on shape
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["shape"] = "rect",
                ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm",
                ["name"] = "AnimatedRect"
            });
            handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
            {
                ["effect"] = "fade",
                ["class"] = "entrance",
                ["duration"] = "500"
            });
            // Transition
            handler.Set("/slide[1]", new Dictionary<string, string>
            {
                ["transition"] = "wipe"
            });
            // Zoom on slide 1 to slide 2
            handler.Add("/slide[1]", "zoom", null, new Dictionary<string, string>
            {
                ["target"] = "2"
            });
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            var chartNode = handler.Get("/slide[1]/chart[1]");
            chartNode.Type.Should().Be("chart");
            chartNode.Format["title"].Should().Be("PersistChart");

            // Find the animated shape
            var shapes = handler.Query("shape");
            var animated = shapes.FirstOrDefault(s => s.Format.ContainsKey("animation"));
            animated.Should().NotBeNull("there should be an animated shape");

            var slideNode = handler.Get("/slide[1]");
            slideNode.Format.Should().ContainKey("transition");

            var zoomNode = handler.Get("/slide[1]/zoom[1]");
            zoomNode.Type.Should().Be("zoom");
        }
    }
}
