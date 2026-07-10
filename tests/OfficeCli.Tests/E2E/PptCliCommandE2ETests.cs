// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using FluentAssertions;

using OfficeCli.Tests.Pptx;

namespace OfficeCli.Tests.E2E;

[Trait("Speed", "E2E")]
public sealed class PptCliCommandE2ETests : PptTestBase
{
    // ==================== Core lifecycle ====================

    [Fact]
    public void Create_OpensAndValidates()
    {
        var path = NewTempPath(".pptx");
        var result = RunCliOk("create", path);
        result.Stdout.Should().Contain("Created");
        File.Exists(path).Should().BeTrue();

        var validate = RunCliOk("validate", path);
        validate.Stdout.Should().Contain("no errors");
    }

    [Fact]
    public void Lifecycle_CreateOpenCloseSaveValidate()
    {
        var path = NewTempPath(".pptx");
        RunCliOk("create", path);
        RunCliOk("open", path);
        RunCliOk("save", path);
        RunCliOk("close", path);
        RunCliOk("validate", path);
    }

    [Fact]
    public void Create_WithForce_OverwritesExisting()
    {
        var path = CreatePresentation();
        var result = RunCliOk("create", path, "--force");
        result.Stderr.Should().Contain("Overwriting");
        File.Exists(path).Should().BeTrue();
        RunCliOk("validate", path);
    }

    [Fact]
    public void Create_WithoutForce_OnExistingFile_Errors()
    {
        var path = CreatePresentation();
        var result = RunCli("create", path);
        result.ExitCode.Should().NotBe(0);
        (result.Stdout + result.Stderr).Should().Contain("already exists");
    }

    // ==================== add ====================

    [Fact]
    public void Add_Slide()
    {
        var path = CreatePresentation();
        var result = RunCliOk("add", path, "/", "--type", "slide");
        result.Stdout.Should().Contain("Added slide").And.Contain("/slide[1]");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("slide");
    }

    [Fact]
    public void Add_SlideWithTitle()
    {
        var path = CreatePresentation();
        var result = RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Hello");
        result.Stdout.Should().Contain("Added slide");
        RunCliOk("validate", path);

        var view = RunCliOk("view", path, "text");
        view.Stdout.Should().Contain("Hello");
    }

    [Fact]
    public void Add_Slide_AtSpecifiedIndex()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=First");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Second");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Third");
        // Insert at index 0 — becomes the first slide
        var result = RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Zero", "--index", "0");
        result.Stdout.Should().Contain("Added slide");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("Zero");
    }

    [Fact]
    public void Add_SlideWithLayout()
    {
        var path = CreatePresentation();
        var result = RunCliOk("add", path, "/", "--type", "slide", "--prop", "layout=blank");
        result.Stdout.Should().Contain("Added slide");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("blank");
    }

    [Fact]
    public void Add_Shape_Rect()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCliOk("add", path, "/slide[1]", "--type", "shape", "--prop", "preset=rect");
        result.Stdout.Should().Contain("Added shape").And.Contain("/slide[1]/shape[");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Add_Shape_WithTextAndPosition()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect",
            "--prop", "text=HelloWorld",
            "--prop", "x=2cm",
            "--prop", "y=3cm",
            "--prop", "width=8cm",
            "--prop", "height=4cm");
        result.Stdout.Should().Contain("Added shape");
        RunCliOk("validate", path);

        var view = RunCliOk("view", path, "text");
        view.Stdout.Should().Contain("HelloWorld");
    }

    [Fact]
    public void Add_Shape_WithBoldAndSize()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect",
            "--prop", "text=BoldText",
            "--prop", "bold=true",
            "--prop", "size=24");
        result.Stdout.Should().Contain("Added shape");
        RunCliOk("validate", path);

        var view = RunCliOk("view", path, "text");
        view.Stdout.Should().Contain("BoldText");
    }

    [Fact]
    public void Add_Shape_WithColor()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect",
            "--prop", "text=ColoredText",
            "--prop", "color=FF0000");
        result.Stdout.Should().Contain("Added shape");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Add_Shape_WithFill()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect",
            "--prop", "text=Filled",
            "--prop", "fill=4472C4");
        result.Stdout.Should().Contain("Added shape");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Add_MultipleShapes_OnSameSlide()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var r1 = RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Shape1");
        r1.Stdout.Should().Contain("/slide[1]/shape[");
        var r2 = RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=ellipse", "--prop", "text=Shape2");
        r2.Stdout.Should().Contain("/slide[1]/shape[");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("shape");
    }

    // ==================== get ====================

    [Fact]
    public void Get_Root_TextOutput()
    {
        var path = CreatePresentation();
        var result = RunCliOk("get", path, "/");
        result.Stdout.Should().Contain("presentation");
        result.Stdout.Should().Contain("slideWidth");
        result.Stdout.Should().Contain("slideHeight");
    }

    [Fact]
    public void Get_Slide_TextOutput()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=TestSlide");
        var result = RunCliOk("get", path, "/slide[1]");
        result.Stdout.Should().Contain("slide").And.Contain("/slide[1]");
    }

    [Fact]
    public void Get_Slide_JsonOutput()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=JsonSlide");
        var result = RunCliOk("get", path, "/slide[1]", "--json");
        result.Stdout.Should().Contain("\"type\"");
        result.Stdout.Should().Contain("\"slide\"");
    }

    [Fact]
    public void Get_Root_JsonOutput()
    {
        var path = CreatePresentation();
        var result = RunCliOk("get", path, "/", "--json");
        result.Stdout.Should().Contain("\"type\"");
        result.Stdout.Should().Contain("\"presentation\"");
    }

    [Fact]
    public void Get_Shape()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var addResult = RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=ShapeText");
        addResult.Stdout.Should().Contain("/slide[1]/shape[");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]/shape[1]");
        getResult.Stdout.Should().Contain("shape");
    }

    // ==================== query ====================

    [Fact]
    public void Query_Slide()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=QuerySlide");
        var result = RunCliOk("query", path, "slide");
        result.Stdout.Should().Contain("/slide[1]");
    }

    [Fact]
    public void Query_Slide_Json()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=QueryJsonSlide");
        var result = RunCliOk("query", path, "slide", "--json");
        result.Stdout.Should().Contain("\"results\"");
        result.Stdout.Should().Contain("/slide[1]");
    }

    [Fact]
    public void Query_Shape()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=FindMe");
        var result = RunCliOk("query", path, "shape");
        result.Stdout.Should().Contain("shape");
    }

    // ==================== set ====================

    [Fact]
    public void Set_SlideHidden()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=ToHide");
        var result = RunCliOk("set", path, "/slide[1]", "--prop", "hidden=true");
        result.Stdout.Should().Contain("Updated");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("hidden");
    }

    [Fact]
    public void Set_Root_SlideSize()
    {
        var path = CreatePresentation();
        var result = RunCliOk("set", path, "/", "--prop", "slideSize=widescreen");
        result.Stdout.Should().Contain("Updated");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/");
        getResult.Stdout.Should().Contain("widescreen");
    }

    [Fact]
    public void Set_Root_DocumentProperties()
    {
        var path = CreatePresentation();
        var result = RunCliOk("set", path, "/",
            "--prop", "title=DocTitle",
            "--prop", "author=DocAuthor",
            "--prop", "subject=DocSubject");
        result.Stdout.Should().Contain("Updated");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/");
        getResult.Stdout.Should().Contain("DocTitle");
    }

    [Fact]
    public void Set_Shape_Text()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Original");
        var result = RunCliOk("set", path, "/slide[1]/shape[1]", "--prop", "text=UpdatedText");
        result.Stdout.Should().Contain("Updated");
        RunCliOk("validate", path);

        var view = RunCliOk("view", path, "text");
        view.Stdout.Should().Contain("UpdatedText");
    }

    [Fact]
    public void Set_Shape_Bold()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=BoldMe");
        var result = RunCliOk("set", path, "/slide[1]/shape[1]", "--prop", "bold=true");
        result.Stdout.Should().Contain("Updated");
        RunCliOk("validate", path);
    }

    // ==================== remove ====================

    [Fact]
    public void Remove_Shape()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=ToDelete");
        var result = RunCliOk("remove", path, "/slide[1]/shape[1]");
        result.Stdout.Should().Contain("Removed");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Remove_Slide()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=ToRemove");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Keep");
        var result = RunCliOk("remove", path, "/slide[1]");
        result.Stdout.Should().Contain("Removed");
        RunCliOk("validate", path);

        // Verify the remaining slide is the one we kept
        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("Keep");
    }

    // ==================== move ====================

    [Fact]
    public void Move_SlideReorder()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=First");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Second");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Third");

        var result = RunCliOk("move", path, "/slide[2]", "--index", "0");
        result.Stdout.Should().Contain("Moved");
        RunCliOk("validate", path);

        // After moving slide[2] to index 0, slide[1] should be the one that was slide[2]
        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("Second");
    }

    // ==================== swap ====================

    [Fact]
    public void Swap_Slides()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Alpha");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Beta");

        var result = RunCliOk("swap", path, "/slide[1]", "/slide[2]");
        result.Stdout.Should().Contain("Swapped");
        RunCliOk("validate", path);

        // After swap, slide[1] should be Beta
        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("Beta");
    }

    // ==================== batch ====================

    [Fact]
    public void Batch_CommandsJson()
    {
        var path = CreatePresentation();
        var commands = """
            [
              {"command":"add","parent":"/","type":"slide","props":{"title":"BatchSlide1"}},
              {"command":"add","parent":"/","type":"slide","props":{"title":"BatchSlide2"}}
            ]
            """;
        var result = RunCliOk("batch", path, "--commands", commands, "--json");
        result.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("BatchSlide1");
    }

    [Fact]
    public void Batch_InputFile()
    {
        var path = CreatePresentation();
        var batchFile = NewTempPath(".json");
        File.WriteAllText(batchFile, """
            [
              {"command":"add","parent":"/","type":"slide","props":{"title":"FileBatch"}},
              {"command":"add","parent":"/slide[1]","type":"shape","props":{"preset":"rect","text":"BatchShape"}}
            ]
            """);

        var result = RunCliOk("batch", path, "--input", batchFile, "--json");
        result.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", path);

        var view = RunCliOk("view", path, "text");
        view.Stdout.Should().Contain("BatchShape");
    }

    [Fact]
    public void Batch_CreateSlideAddShapeSet()
    {
        var path = CreatePresentation();
        var commands = """
            [
              {"command":"add","parent":"/","type":"slide","props":{}},
              {"command":"add","parent":"/slide[1]","type":"shape","props":{"preset":"rect","text":"InitText"}},
              {"command":"set","path":"/slide[1]/shape[1]","props":{"text":"BatchUpdated"}}
            ]
            """;
        var result = RunCliOk("batch", path, "--commands", commands, "--json");
        result.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]/shape[1]");
        getResult.Stdout.Should().Contain("BatchUpdated");

        var view = RunCliOk("view", path, "text");
        view.Stdout.Should().Contain("BatchUpdated");
    }

    // ==================== dump ====================

    [Fact]
    public void Dump_PresentationToStdout()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=DumpTest");
        var result = RunCliOk("dump", path, "/");
        result.Stdout.Should().Contain("\"command\"");
        result.Stdout.Should().Contain("DumpTest");
    }

    [Fact]
    public void Dump_PresentationToFile()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=FileDump");

        var dumpPath = NewTempPath(".json");
        var result = RunCliOk("dump", path, "/", "--out", dumpPath);
        File.Exists(dumpPath).Should().BeTrue();
        var dumpContent = File.ReadAllText(dumpPath);
        dumpContent.Should().Contain("FileDump");
    }

    [Fact]
    public void Dump_ReplayRoundTrip()
    {
        var source = CreatePresentation();
        RunCliOk("add", source, "/", "--type", "slide", "--prop", "title=RoundTripSlide");
        RunCliOk("add", source, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=RTShape");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/", "--out", dumpPath);

        var target = CreatePresentation();
        RunCliOk("batch", target, "--input", dumpPath, "--json");
        RunCliOk("validate", target);

        var view = RunCliOk("view", target, "text");
        view.Stdout.Should().Contain("RTShape");
    }

    // ==================== raw ====================

    [Fact]
    public void Raw_PresentationPart()
    {
        var path = CreatePresentation();
        var result = RunCliOk("raw", path, "/presentation");
        result.Stdout.Should().Contain("<p:presentation");
    }

    [Fact]
    public void Raw_SlidePart()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=RawSlide");
        var result = RunCliOk("raw", path, "/slide[1]");
        result.Stdout.Should().Contain("<p:sld");
    }

    // ==================== raw-set ====================

    [Fact]
    public void RawSet_SetAttribute()
    {
        var path = CreatePresentation();
        var result = RunCliOk("raw-set", path, "/presentation",
            "--xpath", "//p:sldSz",
            "--action", "setattr",
            "--xml", "cx=12192000");
        result.Stdout.Should().Contain("raw-set");
        RunCliOk("validate", path);
    }

    // ==================== add-part ====================

    [Fact]
    public void AddPart_Chart()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCliOk("add-part", path, "/slide[1]", "--type", "chart");
        result.Stdout.Should().Contain("Created chart part");
        // Note: add-part chart creates a skeleton chart part without a chart type;
        // the part is intentionally incomplete until further configuration via set/raw-set.
    }

    // ==================== validate ====================

    [Fact]
    public void Validate_NewPresentation_Passes()
    {
        var path = CreatePresentation();
        var result = RunCliOk("validate", path);
        result.Stdout.Should().Contain("no errors");
    }

    [Fact]
    public void Validate_AfterMutations_Passes()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=ValidSlide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Data");
        RunCliOk("set", path, "/", "--prop", "slideSize=standard");
        RunCliOk("set", path, "/slide[1]/shape[1]", "--prop", "bold=true");

        var result = RunCliOk("validate", path);
        result.Stdout.Should().Contain("no errors");
    }

    // ==================== view ====================

    [Fact]
    public void View_Text_ShowsContent()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=VT1");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=ViewTextContent");

        var result = RunCliOk("view", path, "text");
        result.Stdout.Should().Contain("ViewTextContent");
    }

    [Fact]
    public void View_Text_MultipleSlides()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=SlideA");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=SlideB");

        var result = RunCliOk("view", path, "text");
        result.Stdout.Should().Contain("SlideA");
        result.Stdout.Should().Contain("SlideB");
    }

    [Fact]
    public void View_Outline_ShowsStructure()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=OutlineSlide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=BodyText");

        var result = RunCliOk("view", path, "outline");
        result.Stdout.Should().Contain("OutlineSlide");
    }

    [Fact]
    public void View_Stats_ShowsCounts()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=StatSlide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Hello World");

        var result = RunCliOk("view", path, "stats");
        result.Stdout.Should().Contain("Slides");
        result.Stdout.Should().Contain("shapes");
    }

    [Fact]
    public void View_Issues_ReportsClean()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Titled");

        var result = RunCliOk("view", path, "issues");
        // A properly titled slide should report 0 issues
        result.Stdout.Should().Contain("0 issue");
    }

    // ==================== Error cases ====================

    [Fact]
    public void Error_SetInvalidProperty()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Test");
        var result = RunCli("set", path, "/slide[1]/shape[1]", "--prop", "nonexistent=123");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_AddPictureWithoutSrc()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCli("add", path, "/slide[1]", "--type", "picture");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_RawSetInvalidXPath()
    {
        var path = CreatePresentation();
        var result = RunCli("raw-set", path, "/presentation",
            "--xpath", "///bad[",
            "--action", "setattr",
            "--xml", "foo=bar");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_AddPartUnsupportedType()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCli("add-part", path, "/slide[1]", "--type", "bogus");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_AddPictureInvalidMediaSource()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCli("add", path, "/slide[1]", "--type", "picture",
            "--prop", "src=/nonexistent/file.png");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_DuplicateAddConflict()
    {
        var path = CreatePresentation();
        // Append a second sldSz element — only one is allowed per presentation.
        // raw-set append may succeed but the resulting file is invalid.
        RunCli("raw-set", path, "/presentation",
            "--xpath", "/p:presentation",
            "--action", "append",
            "--xml", "<p:sldSz cx=\"9144000\" cy=\"5143500\" type=\"custom\"/>");
        var validate = RunCli("validate", path);
        validate.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_OpenMissingFile()
    {
        var result = RunCli("open", "nonexistent_file_xyz.pptx");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_GetInvalidSlide()
    {
        var path = CreatePresentation();
        var result = RunCli("get", path, "/slide[99]");
        result.ExitCode.Should().NotBe(0);
        (result.Stdout + result.Stderr).Should().Contain("not found");
    }

    [Fact]
    public void Error_GetInvalidChild()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCli("get", path, "/slide[1]/shape[99]");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_AddInvalidType()
    {
        var path = CreatePresentation();
        var result = RunCli("add", path, "/", "--type", "invalid_type");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_AddShapeWithoutType()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        var result = RunCli("add", path, "/slide[1]");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_RemoveInvalidSlide()
    {
        var path = CreatePresentation();
        var result = RunCli("remove", path, "/slide[0]");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_SetInvalidPath()
    {
        var path = CreatePresentation();
        var result = RunCli("set", path, "/slide[99]", "--prop", "hidden=true");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_ValidateMissingFile()
    {
        var result = RunCli("validate", "nonexistent_file_xyz789.pptx");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Query_NoResults_ReportsWarning()
    {
        var path = CreatePresentation();
        // Querying for shape when there are none should produce output but no matches
        // The "No matches" message is written to stderr by the CLI
        var result = RunCliOk("query", path, "shape");
        result.Stderr.Should().Contain("No matches");
    }

    [Fact]
    public void Error_DumpMissingFile()
    {
        var result = RunCli("dump", "nonexistent_file_xyz_dump.pptx", "/");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_RawSetInvalidPart()
    {
        var path = CreatePresentation();
        var result = RunCli("raw-set", path, "/NoSuchPart",
            "--xpath", "//x:row",
            "--action", "append",
            "--xml", "<x:row r=\"1\"/>");
        result.ExitCode.Should().NotBe(0);
    }

    // ==================== Comprehensive mutation validation ====================

    [Fact]
    public void AllMutations_LeavePresentationValid()
    {
        var path = CreatePresentation();

        // add slides
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Slide 1");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Slide 2");
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "layout=blank");

        // add shapes
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=rect", "--prop", "text=Title Shape",
            "--prop", "bold=true", "--prop", "size=28");
        RunCliOk("add", path, "/slide[1]", "--type", "shape",
            "--prop", "preset=ellipse", "--prop", "text=Body",
            "--prop", "x=2cm", "--prop", "y=4cm");
        RunCliOk("add", path, "/slide[2]", "--type", "shape",
            "--prop", "preset=chevron", "--prop", "text=Slide2Shape",
            "--prop", "fill=4472C4");

        // set properties
        RunCliOk("set", path, "/", "--prop", "title=MultiSlidePres");
        RunCliOk("set", path, "/slide[1]/shape[1]", "--prop", "text=UpdatedTitle");
        RunCliOk("set", path, "/slide[3]", "--prop", "hidden=true");

        // Validate final state
        var validate = RunCliOk("validate", path);
        validate.Stdout.Should().Contain("no errors");
    }

    // ==================== open/close resident flow ====================

    [Fact]
    public void Open_CreatesResidentProcess()
    {
        var path = CreatePresentation();
        var result = RunCliOk("open", path);
        result.Stdout.Should().Contain("Opened");
        RunCliOk("close", path);
    }

    [Fact]
    public void Open_Close_Idempotent()
    {
        var path = CreatePresentation();
        // Close without open should succeed (no resident)
        var result = RunCliOk("close", path);
        result.Stdout.Should().Contain("already saved");
    }

    [Fact]
    public void ResidentFlow_OpenMutateSaveClose_VerifyDiskReadback()
    {
        // 1. Create a blank pptx, open it, add a slide with text, save, verify on disk
        var path = CreatePresentation();
        RunCliOk("open", path);
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=SlideOne");
        RunCliOk("save", path);

        // Verify saved file contains SlideOne (via separate open + view text)
        var afterSavePath = NewTempPath(".pptx");
        File.Copy(path, afterSavePath);
        var view1 = RunCliOk("view", afterSavePath, "text");
        view1.Stdout.Should().Contain("SlideOne");

        // 2. Reopen, add another slide, close, verify on disk
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=SlideTwo");
        RunCliOk("close", path);

        // Verify closed file contains both slides
        var afterClosePath = NewTempPath(".pptx");
        File.Copy(path, afterClosePath);
        var view2 = RunCliOk("view", afterClosePath, "text");
        view2.Stdout.Should().Contain("SlideOne");
        view2.Stdout.Should().Contain("SlideTwo");
    }

    // ==================== view edge cases ====================

    [Fact]
    public void View_Text_UntitledSlide()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");

        var result = RunCliOk("view", path, "text");
        // Slides without titles show "(untitled)" or similar indication
        result.Stdout.Should().Contain("/slide[1]");
    }

    [Fact]
    public void View_Stats_EmptyPresentation()
    {
        var path = CreatePresentation();
        var result = RunCliOk("view", path, "stats");
        result.Stdout.Should().Contain("Slides");
    }

    // ==================== get edge cases ====================

    [Fact]
    public void Get_Slide_AfterSet()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide");
        RunCliOk("set", path, "/slide[1]", "--prop", "hidden=true");

        var result = RunCliOk("get", path, "/slide[1]");
        result.Stdout.Should().Contain("hidden");
    }

    // ==================== removal + re-add ====================

    [Fact]
    public void RemoveThenReAdd_Slide()
    {
        var path = CreatePresentation();
        RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Gone");
        RunCliOk("remove", path, "/slide[1]");
        var result = RunCliOk("add", path, "/", "--type", "slide", "--prop", "title=Back");
        result.Stdout.Should().Contain("/slide[1]");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/slide[1]");
        getResult.Stdout.Should().Contain("Back");
    }
}
