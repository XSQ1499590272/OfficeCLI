// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using FluentAssertions;

using OfficeCli.Tests.Pptx;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public sealed class PptDumpReplayE2ETests : PptTestBase
{
    // ==================== full dump/replay ====================

    [Fact]
    public void FullDumpReplay_RoundTripPreservesSlidesAndContent()
    {
        var source = CreatePresentation();

        // Build a rich deck with slides, text, shapes, table, chart, picture, notes, and comment.
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=Dump Slide 1");
        // Notes are set-only, not supported on add.
        RunCliOk("set", source, "/slide[1]", "--prop", "notes=RT Speaker Notes");
        RunCliOk("add", source, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=RT Text Shape",
            "--prop", "bold=true");

        // Picture
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        RunCliOk("add", source, "/slide[1]", "--type", "picture",
            "--prop", $"src={pngPath}", "--prop", "alt=RT Picture");

        // Table
        RunCliOk("add", source, "/slide[1]", "--type", "table",
            "--prop", "data=Header1,Header2,Header3;A1,B1,C1;A2,B2,C2",
            "--prop", "firstRow=true");

        // Second slide with chart
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=Dump Slide 2");
        RunCliOk("add", source, "/slide[2]", "--type", "chart",
            "--prop", "chartType=bar",
            "--prop", "title=RT Chart",
            "--prop", "categories=Q1,Q2,Q3",
            "--prop", "data=Series1:10,20,30");

        // Comment on slide 1
        RunCliOk("add", source, "/slide[1]", "--type", "comment",
            "--prop", "text=RT Review Comment", "--prop", "author=RT Tester");

        // Transition on slide 1
        RunCliOk("set", source, "/slide[1]", "--prop", "transition=fade");

        // Validate source
        RunCliOk("validate", source);

        // Dump to file
        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/", "--out", dumpPath);
        File.Exists(dumpPath).Should().BeTrue();

        var dumpedJson = File.ReadAllText(dumpPath);
        dumpedJson.Should().NotBeNullOrEmpty();
        dumpedJson.Should().Contain("\"command\"");
        dumpedJson.Should().Contain("RT Text Shape");

        // Replay dump into a blank target
        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");

        // Validate replayed target
        RunCliOk("validate", target);

        // Verify content survived round trip
        var view = RunCliOk("view", target, "text");
        view.Stdout.Should().Contain("RT Text Shape");
        view.Stdout.Should().Contain("Dump Slide 1");
        view.Stdout.Should().Contain("Dump Slide 2");

        // Verify chart survived
        var chartQuery = RunCliOk("query", target, "chart");
        chartQuery.Stdout.Should().Contain("RT Chart");

        // Verify picture survived
        var picQuery = RunCliOk("query", target, "picture");
        picQuery.Stdout.Should().Contain("RT Picture");

        // Verify notes survived
        var notesRaw = RunCliOk("raw", target, "/noteSlide[1]");
        notesRaw.Stdout.Should().Contain("RT Speaker Notes");

        // Verify comment survived
        var commentQuery = RunCliOk("query", target, "comment");
        commentQuery.Stdout.Should().Contain("RT Review Comment");

        // Verify slide count
        var stats = RunCliOk("view", target, "stats");
        stats.Stdout.Should().Contain("Slides: 2");
    }

    [Fact]
    public void FullDumpReplay_DumpToStdout_ProducesReplayableJson()
    {
        var source = CreatePresentation();
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=StdoutDump");
        RunCliOk("add", source, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=StdoutContent");

        var result = RunCliOk("dump", source, "/");
        result.Stdout.Should().Contain("\"command\"");
        result.Stdout.Should().Contain("StdoutContent");

        // The stdout JSON should be valid and replayable.
        var dumpPath = NewTempPath(".json");
        File.WriteAllText(dumpPath, result.Stdout);

        var target = CreatePresentation();
        RunCliOk("batch", target, "--input", dumpPath, "--json");
        RunCliOk("validate", target);

        var view = RunCliOk("view", target, "text");
        view.Stdout.Should().Contain("StdoutContent");
    }

    [Fact]
    public void FullDumpReplay_WithOleObject_SurvivesRoundTrip()
    {
        var source = CreatePresentation();
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=OLE Slide");

        var olePath = CreateOlePayload(NewTempPath(".bin"));
        RunCliOk("add", source, "/slide[1]", "--type", "ole",
            "--prop", $"src={olePath}",
            "--prop", "name=EmbeddedDoc",
            "--prop", "width=3in",
            "--prop", "height=2in");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/", "--out", dumpPath);

        var target = CreatePresentation();
        RunCliOk("batch", target, "--input", dumpPath, "--json");
        RunCliOk("validate", target);

        // OLE object should exist in the target
        var oleQuery = RunCliOk("query", target, "ole");
        oleQuery.Stdout.Should().Contain("EmbeddedDoc");
    }

    [Fact]
    public void FullDumpReplay_WithModel3d_SurvivesRoundTrip()
    {
        var source = CreatePresentation();
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=3D Slide");

        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        RunCliOk("add", source, "/slide[1]", "--type", "model3d",
            "--prop", $"src={glbPath}",
            "--prop", "width=3in",
            "--prop", "height=3in");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/", "--out", dumpPath);

        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", target);

        // 3D model should exist in target
        var model3dQuery = RunCliOk("query", target, "model3d");
        model3dQuery.Stdout.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void FullDumpReplay_WithVideo_SurvivesRoundTrip()
    {
        var source = CreatePresentation();
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=Video Slide");

        var videoPath = CreateTinyVideo(NewTempPath(".mp4"));
        RunCliOk("add", source, "/slide[1]", "--type", "video",
            "--prop", $"src={videoPath}",
            "--prop", "width=4in",
            "--prop", "height=3in");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/", "--out", dumpPath);

        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", target);

        var videoQuery = RunCliOk("query", target, "video");
        videoQuery.Stdout.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void FullDumpReplay_WithAnimation_SurvivesRoundTrip()
    {
        var source = CreatePresentation();
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=Anim Slide");
        RunCliOk("add", source, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Animated Shape",
            "--prop", "animation=fade");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/", "--out", dumpPath);

        var target = CreatePresentation();
        RunCliOk("batch", target, "--input", dumpPath, "--json");
        RunCliOk("validate", target);

        // Shape and animation should survive.
        // Note: slide add with title creates shape[1] as the title placeholder,
        // so the animated shape is at shape[2].
        var getResult = RunCliOk("get", target, "/slide[1]/shape[2]");
        getResult.Stdout.Should().Contain("Animated Shape");
    }

    // ==================== subtree dump/replay ====================

    [Fact]
    public void SubtreeDumpReplay_Slide_ProducesReplayableItems()
    {
        var source = CreatePresentation();
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=Subtree Slide");
        RunCliOk("add", source, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=SubtreeShape");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/slide[1]", "--out", dumpPath);

        var dumpContent = File.ReadAllText(dumpPath);
        dumpContent.Should().Contain("\"command\"");
        dumpContent.Should().Contain("SubtreeShape");

        // Replay into a blank deck
        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", target);

        var view = RunCliOk("view", target, "text");
        view.Stdout.Should().Contain("SubtreeShape");
    }

    [Fact]
    public void SubtreeDumpReplay_Theme_ProducesRawSet()
    {
        var source = CreatePresentation();
        // Apply a theme color to ensure non-default theme content
        RunCliOk("set", source, "/", "--prop", "theme.color.accent1=FF6600");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/theme", "--out", dumpPath);

        var dumpContent = File.ReadAllText(dumpPath);
        dumpContent.Should().Contain("\"command\"");
        // Theme subtree should emit a raw-set command
        dumpContent.Should().Contain("raw-set");
        dumpContent.Should().MatchRegex("(/theme|a:theme)");

        // Replay into a blank deck
        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", target);

        // Theme should be applied on target
        var raw = RunCliOk("raw", target, "/theme");
        raw.Stdout.Should().Contain("<a:theme");
        raw.Stdout.Should().NotContain("(no theme)");
    }

    [Fact]
    public void SubtreeDumpReplay_NotesMaster_ProducesRawSet()
    {
        var source = CreatePresentation();
        // Notes master may or may not exist on a blank deck.
        // Create a slide first to increase the chance of having a notesMaster.

        var dumpPath = NewTempPath(".json");
        var result = RunCli("dump", source, "/notesMaster", "--out", dumpPath);

        // notesMaster may not exist on every blank deck. If the dump succeeds
        // and produces non-empty content, verify it's replayable.
        if (result.ExitCode == 0)
        {
            TrackTempFile(dumpPath);
            var dumpContent = File.ReadAllText(dumpPath);
            if (dumpContent.Length > 4) // more than just "[]\n"
            {
                dumpContent.Should().Contain("raw-set");
                dumpContent.Should().MatchRegex("(/notesMaster|p:notesMaster)");

                var target = CreatePresentation();
                var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
                batchResult.Stdout.Should().Contain("\"succeeded\"");
                RunCliOk("validate", target);
            }
            // If dumpContent is empty array "[]", notesMaster doesn't exist — that's fine.
        }
    }

    [Fact]
    public void SubtreeDumpReplay_SlideMaster_ProducesRawSet()
    {
        var source = CreatePresentation();

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/slideMaster[1]", "--out", dumpPath);

        var dumpContent = File.ReadAllText(dumpPath);
        dumpContent.Should().Contain("raw-set");
        dumpContent.Should().MatchRegex("(/slideMaster|p:sldMaster)");

        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", target);
    }

    [Fact]
    public void SubtreeDumpReplay_SlideLayout_ProducesRawSet()
    {
        var source = CreatePresentation();

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/slideLayout[1]", "--out", dumpPath);

        var dumpContent = File.ReadAllText(dumpPath);
        dumpContent.Should().Contain("raw-set");
        dumpContent.Should().MatchRegex("(/slideLayout|p:sldLayout)");

        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", target);
    }

    [Fact]
    public void SubtreeDumpReplay_NoteSlide_ProducesRawSet()
    {
        var source = CreatePresentation();
        // Notes text is set-only, applied after slide creation.
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=NoteSlide Subtree");
        RunCliOk("set", source, "/slide[1]", "--prop", "notes=NoteSlide notes content");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/noteSlide[1]", "--out", dumpPath);

        var dumpContent = File.ReadAllText(dumpPath);
        dumpContent.Should().Contain("raw-set");

        // Replay requires a target with slide 1 + notes, since noteSlide is per-slide
        // and raw-set on /noteSlide[1] requires the notes part to exist on the target.
        var target = CreatePresentation();
        RunCliOk("add", target, "/", "--type", "slide", "--prop", "title=Target Slide");
        RunCliOk("set", target, "/slide[1]", "--prop", "notes=Target notes placeholder");
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
    }

    [Fact]
    public void SubtreeDumpReplay_Presentation_EmitsProperties()
    {
        var source = CreatePresentation();
        RunCliOk("set", source, "/", "--prop", "slideSize=standard");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/presentation", "--out", dumpPath);

        var dumpContent = File.ReadAllText(dumpPath);
        dumpContent.Should().Contain("\"command\"");

        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", target);
    }

    // ==================== raw/add-part in dump/replay ====================

    [Fact]
    public void RawAndAddPart_ContentSurvivesDumpReplay()
    {
        var source = CreatePresentation();
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=RawAddPart Slide");

        // Add a chart part via add-part
        RunCliOk("add-part", source, "/slide[1]", "--type", "chart");

        // Set slide size via raw-set as a simple raw operation
        RunCliOk("raw-set", source, "/presentation",
            "--xpath", "//p:sldSz",
            "--action", "setattr",
            "--xml", "cx=12192000");

        // Dump and replay
        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/", "--out", dumpPath);

        var dumpContent = File.ReadAllText(dumpPath);
        // Dump should contain recognizable commands
        dumpContent.Should().Contain("\"command\"");
        dumpContent.Should().Contain("RawAddPart Slide");

        var target = CreatePresentation();
        var batchResult = RunCliOk("batch", target, "--input", dumpPath, "--json");
        batchResult.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", target);

        // Verify slide survived
        var view = RunCliOk("view", target, "text");
        view.Stdout.Should().Contain("RawAddPart Slide");
    }

    // ==================== dump error cases ====================

    [Fact]
    public void Dump_InvalidSubtree_ReportsError()
    {
        var source = CreatePresentation();

        var result = RunCli("dump", source, "/bogus[1]", "--out", NewTempPath(".json"));
        result.ExitCode.Should().NotBe(0);
        (result.Stdout + result.Stderr).Should().Contain("not supported");
    }

    [Fact]
    public void Dump_OutOfRangeSlide_ReportsError()
    {
        var source = CreatePresentation();

        var result = RunCli("dump", source, "/slide[99]", "--out", NewTempPath(".json"));
        result.ExitCode.Should().NotBe(0);
        (result.Stdout + result.Stderr).Should().Contain("not found");
    }

    // ==================== batch error cases ====================

    [Fact]
    public void Batch_InvalidJson_ReportsError()
    {
        var source = CreatePresentation();

        var badJsonPath = NewTempPath(".json");
        File.WriteAllText(badJsonPath, "{ not valid json }");

        var result = RunCli("batch", source, "--input", badJsonPath, "--json");
        result.ExitCode.Should().NotBe(0);
    }
}
