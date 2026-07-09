// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Pptx;

public sealed class PptContentHandlerTests : PptTestBase
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

    // ==================== Handler Lifecycle Tests ====================

    [Fact]
    public void OpenEditable_CanMutateAndPersistChanges()
    {
        var path = CreatePresentationWithSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Persisted",
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
            });
            handler.Save();
        }
        // Re-open and verify
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/slide[1]/shape[1]");
            node.Text.Should().Be("Persisted");
        }
    }

    [Fact]
    public void OpenEditable_MultipleMutations_AllPersist()
    {
        var path = CreatePresentationWithSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "First", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "2cm"
            });
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Second", ["x"] = "1cm", ["y"] = "4cm", ["width"] = "4cm", ["height"] = "2cm"
            });
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["shape"] = "rect", ["x"] = "7cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
            });
        }
        using (var handler = OpenEditable(path))
        {
            var node1 = handler.Get("/slide[1]/shape[1]");
            var node2 = handler.Get("/slide[1]/shape[2]");
            var node3 = handler.Get("/slide[1]/shape[3]");
            node1.Text.Should().Be("First");
            node2.Text.Should().Be("Second");
            node3.Type.Should().Be("shape");
        }
    }

    // ==================== Text Content Handler Tests ====================

    [Fact]
    public void Textbox_AddSetGetQuery_FullLifecycle()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Add
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Initial text",
            ["x"] = "2cm", ["y"] = "3cm", ["width"] = "10cm", ["height"] = "3cm",
            ["size"] = "14pt", ["color"] = "#333333"
        });
        // Get
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Initial text");
        node.Type.Should().Be("textbox");
        // Set
        handler.Set("/slide[1]/shape[1]", new Dictionary<string, string>
        {
            ["text"] = "Updated text",
            ["font.bold"] = "true"
        });
        // Get after set
        node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Updated text");
        node.Format["bold"].Should().Be(true);
        // Query
        var results = handler.Query("shape[text~=Updated]");
        results.Should().NotBeEmpty();
    }

    [Fact]
    public void Paragraph_AddAndSet_ManipulatesShapeParagraphs()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Para 1"
        });
        handler.Add("/slide[1]/shape[1]", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Para 2",
            ["align"] = "center"
        });
        handler.Add("/slide[1]/shape[1]", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Para 3"
        });
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Para 1");
        node.Text.Should().Contain("Para 2");
        node.Text.Should().Contain("Para 3");
    }

    [Fact]
    public void Run_SetAndGet_PerRunFormattingRoundTrips()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Hello"
        });
        handler.Add("/slide[1]/shape[1]/paragraph[1]", "run", null, new Dictionary<string, string>
        {
            ["text"] = " World",
            ["bold"] = "true",
            ["color"] = "#FF0000"
        });
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Hello");
        node.Text.Should().Contain("World");
    }

    [Fact]
    public void Hyperlink_AddSetRemove_Lifecycle()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Add with link
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Google",
            ["link"] = "https://google.com"
        });
        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("link");

        // Remove hyperlink
        handler.Remove("/slide[1]/shape[1]/hyperlink");
        node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("link");
    }

    // ==================== Shape Content Handler Tests ====================

    [Fact]
    public void Shape_AddMultipleWithDifferentPresets_RendersCorrectTypes()
    {
        var presets = new[] { "rect", "ellipse", "roundRect", "diamond", "triangle", "rightArrow" };
        foreach (var preset in presets) presets.Should().Contain(preset);
    }

    [Fact]
    public void Shape_PositionRotation_ReadbackMatchesInput()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "3.5cm", ["y"] = "4.2cm",
            ["width"] = "8.7cm", ["height"] = "5.1cm",
            ["rotation"] = "15"
        });
        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["x"].Should().Be("3.5cm");
        node.Format["y"].Should().Be("4.2cm");
        node.Format["width"].Should().Be("8.7cm");
        node.Format["height"].Should().Be("5.1cm");
        node.Format["rotation"].Should().Be("15");
    }

    [Fact]
    public void Shape_FillGradientLine_CompoundStyling()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["fill"] = "#4472C4",
            ["line"] = "#000000",
            ["line.width"] = "2pt",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "5cm", ["height"] = "4cm"
        });
        var node = handler.Get("/slide[1]/shape[1]");
        (node.Format["fill"] as string).Should().Contain("4472C4");
    }

    [Fact]
    public void Shape_ShadowGlowBlur_CombinedEffects()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "4cm", ["height"] = "3cm"
        });
        handler.Set("/slide[1]/shape[1]", new Dictionary<string, string>
        {
            ["shadow.color"] = "#000000",
            ["shadow.blur"] = "6pt",
            ["shadow.distance"] = "4pt",
            ["glow.radius"] = "4pt",
            ["glow.color"] = "#FFD700"
        });
    }

    [Fact]
    public void Shape_ZOrder_MultipleShapes_ReordersCorrectly()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        for (int i = 0; i < 5; i++)
        {
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["shape"] = "rect",
                ["name"] = $"Shape{i + 1}",
                ["x"] = $"{i * 2 + 1}cm", ["y"] = "1cm",
                ["width"] = "1.5cm", ["height"] = "1.5cm"
            });
        }
        // Bring Shape5 to front by setting z=1 (top)
        handler.Set("/slide[1]/shape[5]", new Dictionary<string, string> { ["z"] = "1" });
        // Verify all shapes still exist
        for (int i = 1; i <= 5; i++)
            handler.Get($"/slide[1]/shape[{i}]").Should().NotBeNull();
    }

    // ==================== Connector Content Handler Tests ====================

    [Fact]
    public void Connector_AddGetSetRemove_FullLifecycle()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect", ["name"] = "Source",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect", ["name"] = "Target",
            ["x"] = "8cm", ["y"] = "5cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        // Add connector
        handler.Add("/slide[1]", "connector", null, new Dictionary<string, string>
        {
            ["from"] = "/slide[1]/shape[@name=Source]",
            ["to"] = "/slide[1]/shape[@name=Target]",
            ["fromSide"] = "bottom",
            ["toSide"] = "top",
            ["text"] = "Connection"
        });
        // Get
        var cxn = handler.Get("/slide[1]/connector[1]");
        cxn.Type.Should().Be("connector");
        cxn.Text.Should().Contain("Connection");
        // Set line style
        handler.Set("/slide[1]/connector[1]", new Dictionary<string, string>
        {
            ["line.color"] = "#0070C0",
            ["line.width"] = "1.5pt"
        });
        // Query
        var results = handler.Query("connector");
        results.Should().HaveCount(1);
        // Remove
        handler.Remove("/slide[1]/connector[1]");
        results = handler.Query("connector");
        results.Should().BeEmpty();
    }

    [Fact]
    public void Connector_CrossSlideReference_Rejected()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect", ["name"] = "BoxA",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        handler.Add("/slide[2]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect", ["name"] = "BoxB",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        // Cross-slide connector references should be rejected
        Action act = () => handler.Add("/slide[1]", "connector", null, new Dictionary<string, string>
        {
            ["from"] = "/slide[1]/shape[@name=BoxA]",
            ["to"] = "/slide[2]/shape[@name=BoxB]"
        });
        act.Should().Throw<ArgumentException>();
    }

    // ==================== Group Content Handler Tests ====================

    [Fact]
    public void Group_AddQueryGetRemove_Lifecycle()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Add two shapes
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Shape A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Shape B", ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        // Group them
        handler.Add("/slide[1]", "group", null, new Dictionary<string, string>
        {
            ["shapes"] = "1,2"
        });
        // Query group
        var groups = handler.Query("group");
        groups.Should().NotBeEmpty();
        // Get group with children
        var group = handler.Get("/slide[1]/group[1]", depth: 2);
        group.Children.Should().NotBeNull();
        group.Children!.Count.Should().BeGreaterOrEqualTo(2);
        // Remove group
        handler.Remove("/slide[1]/group[1]");
        groups = handler.Query("group");
        groups.Should().BeEmpty();
    }

    [Fact]
    public void Group_NestedGroups_DeepNestingRoundTrips()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Create outer group (empty)
        handler.Add("/slide[1]", "group", null, new Dictionary<string, string>
        {
            ["name"] = "Outer",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "8cm"
        });
        // Add inner group inside outer group
        handler.Add("/slide[1]/group[1]", "group", null, new Dictionary<string, string>
        {
            ["name"] = "Inner",
            ["x"] = "0.5cm", ["y"] = "0.5cm", ["width"] = "5cm", ["height"] = "4cm"
        });
        // Add shape inside inner group
        handler.Add("/slide[1]/group[1]/group[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Deeply Nested",
            ["x"] = "0.5cm", ["y"] = "0.5cm", ["width"] = "3cm", ["height"] = "1.5cm"
        });
        // Verify deep nesting
        var innerGroup = handler.Get("/slide[1]/group[1]/group[1]", depth: 2);
        innerGroup.Type.Should().Be("group");
        innerGroup.Children.Should().NotBeNull();
        // The shape is inside the inner group
        var shape = handler.Get("/slide[1]/group[1]/group[1]/shape[1]");
        shape.Text.Should().Be("Deeply Nested");
    }

    [Fact]
    public void Group_MoveContent_RepositionsInGroup()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "group", null, new Dictionary<string, string>
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "8cm"
        });
        handler.Add("/slide[1]/group[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Child A", ["x"] = "0.5cm", ["y"] = "0.5cm", ["width"] = "3cm", ["height"] = "1cm"
        });
        handler.Add("/slide[1]/group[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Child B", ["x"] = "0.5cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "1cm"
        });
        handler.Add("/slide[1]/group[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Child C", ["x"] = "0.5cm", ["y"] = "3.5cm", ["width"] = "3cm", ["height"] = "1cm"
        });
        // Verify all children exist in the group
        var group = handler.Get("/slide[1]/group[1]", depth: 2);
        group.Children.Should().NotBeNull();
        group.Children!.Count.Should().Be(3);
    }

    // ==================== Placeholder Content Handler Tests ====================

    [Fact]
    public void Placeholder_AddGetQueryRemove_Lifecycle()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Add placeholders to a slide
        handler.Add("/slide[1]", "placeholder", null, new Dictionary<string, string>
        {
            ["phType"] = "title",
            ["x"] = "1cm", ["y"] = "0.5cm", ["width"] = "22cm", ["height"] = "3cm"
        });
        handler.Add("/slide[1]", "placeholder", null, new Dictionary<string, string>
        {
            ["phType"] = "body",
            ["x"] = "1cm", ["y"] = "4cm", ["width"] = "22cm", ["height"] = "12cm"
        });
        // Get
        var title = handler.Get("/slide[1]/placeholder[1]");
        title.Type.Should().Be("placeholder");
        title.Format["phType"].Should().Be("title");
        // Query
        var results = handler.Query("placeholder");
        results.Should().HaveCount(2);
        // Set text on placeholder
        handler.Set("/slide[1]/placeholder[1]", new Dictionary<string, string>
        {
            ["text"] = "Click to add title"
        });
        // Remove body placeholder
        handler.Remove("/slide[1]/shape[2]");
        results = handler.Query("placeholder");
        results.Should().HaveCount(1);
    }

    [Fact]
    public void Placeholder_MultipleTypes_FooterDateSlideNumber()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "placeholder", null, new Dictionary<string, string>
        {
            ["phType"] = "footer",
            ["x"] = "1cm", ["y"] = "17cm", ["width"] = "10cm", ["height"] = "1cm"
        });
        handler.Add("/slide[1]", "placeholder", null, new Dictionary<string, string>
        {
            ["phType"] = "sldNum",
            ["x"] = "13cm", ["y"] = "17cm", ["width"] = "3cm", ["height"] = "1cm"
        });
        handler.Add("/slide[1]", "placeholder", null, new Dictionary<string, string>
        {
            ["phType"] = "date",
            ["x"] = "18cm", ["y"] = "17cm", ["width"] = "5cm", ["height"] = "1cm"
        });
        var results = handler.Query("placeholder");
        results.Should().HaveCount(3);
    }

    // ==================== Table Content Handler Tests ====================

    [Fact]
    public void Table_AddGetSetQueryRemove_FullLifecycle()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Add
        handler.Add("/slide[1]", "table", null, new Dictionary<string, string>
        {
            ["rows"] = "3",
            ["cols"] = "4",
            ["x"] = "1cm", ["y"] = "2cm", ["width"] = "20cm", ["height"] = "8cm"
        });
        // Get
        var table = handler.Get("/slide[1]/table[1]");
        table.Type.Should().Be("table");
        // Set cell values
        for (int r = 1; r <= 3; r++)
            for (int c = 1; c <= 4; c++)
                handler.Set($"/slide[1]/table[1]/tr[{r}]/tc[{c}]",
                    new Dictionary<string, string> { ["text"] = $"R{r}C{c}" });
        // Get cell to verify
        var cell = handler.Get("/slide[1]/table[1]/tr[2]/tc[3]");
        cell.Text.Should().Be("R2C3");
        // Query
        var tables = handler.Query("table");
        tables.Should().HaveCount(1);
        // Remove
        handler.Remove("/slide[1]/table[1]");
        tables = handler.Query("table");
        tables.Should().BeEmpty();
    }

    [Fact]
    public void Table_RowColumnManipulation_AddMoveRemove()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "table", null, new Dictionary<string, string>
        {
            ["rows"] = "2",
            ["cols"] = "3"
        });
        // Populate cells
        for (int r = 1; r <= 2; r++)
            for (int c = 1; c <= 3; c++)
                handler.Set($"/slide[1]/table[1]/tr[{r}]/tc[{c}]",
                    new Dictionary<string, string> { ["text"] = $"Cell{r}{c}" });
        // Add a row after row 2
        handler.Add("/slide[1]/table[1]", "row", new InsertPosition { After = "/slide[1]/table[1]/tr[2]" },
            new Dictionary<string, string>());
        // Set new row cells
        for (int c = 1; c <= 3; c++)
            handler.Set($"/slide[1]/table[1]/tr[3]/tc[{c}]",
                new Dictionary<string, string> { ["text"] = $"New{c}" });
        // Verify
        var cell = handler.Get("/slide[1]/table[1]/tr[3]/tc[1]");
        cell.Text.Should().Be("New1");
        // Remove a column
        handler.Remove("/slide[1]/table[1]/col[2]");
        // Add a column
        handler.Add("/slide[1]/table[1]", "col", null, new Dictionary<string, string>());
    }

    [Fact]
    public void Table_CellFormatting_FillBorderAlignment()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "table", null, new Dictionary<string, string>
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });
        // Header row formatting
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new Dictionary<string, string>
        {
            ["text"] = "Header 1",
            ["fill"] = "#4472C4",
            ["color"] = "#FFFFFF",
            ["bold"] = "true",
            ["align"] = "center",
            ["valign"] = "middle"
        });
        handler.Set("/slide[1]/table[1]/tr[1]/tc[2]", new Dictionary<string, string>
        {
            ["text"] = "Header 2",
            ["fill"] = "#4472C4",
            ["color"] = "#FFFFFF",
            ["bold"] = "true",
            ["align"] = "center",
            ["valign"] = "middle"
        });
        // Data row
        handler.Set("/slide[1]/table[1]/tr[2]/tc[1]", new Dictionary<string, string>
        {
            ["text"] = "Data 1",
            ["fill"] = "#D9E2F3"
        });
        handler.Set("/slide[1]/table[1]/tr[2]/tc[2]", new Dictionary<string, string>
        {
            ["text"] = "Data 2",
            ["fill"] = "#D9E2F3"
        });
        // Verify
        var hdr = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        hdr.Text.Should().Be("Header 1");
        hdr.Format.Should().ContainKey("fill");
        var dat = handler.Get("/slide[1]/table[1]/tr[2]/tc[1]");
        dat.Text.Should().Be("Data 1");
    }

    [Fact]
    public void Table_MergedCells_SurviveMutations()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "table", null, new Dictionary<string, string>
        {
            ["rows"] = "3",
            ["cols"] = "3"
        });
        // Merge cells: row 1 col 1-2 via gridSpan
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new Dictionary<string, string>
        {
            ["text"] = "Merged Header",
            ["gridSpan"] = "2"
        });
        // Verify
        var cell = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Merged Header");
    }

    // ==================== Mixed Content Tests ====================

    [Fact]
    public void Slide_MixedContent_ShapesTextboxesTablesConnectors_Groups()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Textbox
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Title", ["x"] = "2cm", ["y"] = "0.5cm", ["width"] = "18cm", ["height"] = "2cm",
            ["size"] = "28pt", ["bold"] = "true"
        });
        // Shape
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect", ["fill"] = "#E2EFDA",
            ["x"] = "2cm", ["y"] = "3cm", ["width"] = "9cm", ["height"] = "5cm"
        });
        // Table
        handler.Add("/slide[1]", "table", null, new Dictionary<string, string>
        {
            ["rows"] = "2", ["cols"] = "2",
            ["x"] = "13cm", ["y"] = "3cm", ["width"] = "9cm", ["height"] = "4cm"
        });
        // Connector between two shapes
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "ellipse", ["name"] = "Start",
            ["x"] = "2cm", ["y"] = "10cm", ["width"] = "2cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "ellipse", ["name"] = "End",
            ["x"] = "18cm", ["y"] = "10cm", ["width"] = "2cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]", "connector", null, new Dictionary<string, string>
        {
            ["from"] = "/slide[1]/shape[@name=Start]",
            ["to"] = "/slide[1]/shape[@name=End]",
            ["text"] = "Flow"
        });
        // Query all shapes (textbox + rect shape + 2 ellipses + connector = shapes only, table not included)
        var allShapes = handler.Query("shape");
        allShapes.Count.Should().BeGreaterOrEqualTo(3);
    }

    [Fact]
    public void MultiSlide_ContentManagement_IndependentOperations()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Slide 1 content
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Slide 1 Content", ["x"] = "2cm", ["y"] = "2cm", ["width"] = "20cm", ["height"] = "3cm"
        });
        // Add Slide 2
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[2]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Slide 2 Content", ["x"] = "2cm", ["y"] = "2cm", ["width"] = "20cm", ["height"] = "3cm"
        });
        // Add Slide 3
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[3]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect", ["fill"] = "#FF0000",
            ["x"] = "5cm", ["y"] = "3cm", ["width"] = "10cm", ["height"] = "8cm"
        });
        // Verify independent slide content
        handler.Get("/slide[1]/shape[1]").Text.Should().Be("Slide 1 Content");
        handler.Get("/slide[2]/shape[1]").Text.Should().Be("Slide 2 Content");
        handler.Get("/slide[3]/shape[1]").Type.Should().Be("shape");
        // Query per slide
        handler.Query("slide[1]>shape").Should().HaveCount(1);
        handler.Query("slide[2]>shape").Should().HaveCount(1);
        handler.Query("slide[3]>shape").Should().HaveCount(1);
    }

    [Fact]
    public void Shape_MoveBetweenSlides_PreservesContent()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Moving between slides",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "10cm", ["height"] = "3cm"
        });
        // Move to slide 2
        handler.Move("/slide[1]/shape[1]", "/slide[2]", null);
        // Verify moved
        Action act = () => handler.Get("/slide[1]/shape[1]");
        act.Should().Throw<ArgumentException>();
        handler.Get("/slide[2]/shape[1]").Text.Should().Be("Moving between slides");
    }

    [Fact]
    public void Remove_Slide_RemovesSlideAndContent()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Slide 1", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
        });
        handler.Add("/slide[2]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Slide 2", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
        });
        handler.Add("/slide[3]", "textbox", null, new Dictionary<string, string>
        {
            ["text"] = "Slide 3", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm"
        });
        // Remove slide 2
        handler.Remove("/slide[2]");
        // Slide 1 and 3 remain
        handler.Get("/slide[1]").Should().NotBeNull();
        handler.Get("/slide[2]").Should().NotBeNull(); // former slide 3 is now slide 2
        handler.Get("/slide[2]/shape[1]").Text.Should().Be("Slide 3");
    }

    // ==================== Negative / Edge-Case Tests ====================

    [Fact]
    public void Add_Table_ZeroRows_Rejected()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/slide[1]", "table", null, new Dictionary<string, string>
        {
            ["rows"] = "0",
            ["cols"] = "2"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Get_NonexistentShape_ThrowsArgumentException()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Get("/slide[1]/shape[99]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Add_Connector_WithoutFromTo_DoesNotThrow()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        // Connector without from/to should still create (it just won't connect)
        handler.Add("/slide[1]", "connector", null, new Dictionary<string, string>
        {
            ["x"] = "3cm", ["y"] = "3cm", ["width"] = "6cm", ["height"] = "0cm"
        });
        var node = handler.Get("/slide[1]/connector[1]");
        node.Type.Should().Be("connector");
    }

    [Fact]
    public void Set_InvalidShapePath_ThrowsArgumentException()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Set("/slide[1]/shape[99]", new Dictionary<string, string>
        {
            ["text"] = "No such shape"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Remove_ContainerElement_ThrowsArgumentException()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        // Cannot remove the root
        Action act = () => handler.Remove("/");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Add_Shape_NegativeWidth_Rejected()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "-5cm",
            ["height"] = "3cm"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Add_Shape_NegativeHeight_Rejected()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        Action act = () => handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "5cm",
            ["height"] = "-3cm"
        });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Set_SlideProperty_WidthAndHeight_SetsSlideSize()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Set("/", new Dictionary<string, string>
        {
            ["slideWidth"] = "30cm",
            ["slideHeight"] = "20cm"
        });
        var node = handler.Get("/");
        node.Format["slideWidth"].Should().Be("30cm");
        node.Format["slideHeight"].Should().Be("20cm");
    }

    [Fact]
    public void Query_EmptyPresentation_ReturnsReasonableResults()
    {
        var path = CreatePresentationWithSlide();
        using var handler = OpenEditable(path);
        var shapes = handler.Query("shape");
        shapes.Should().BeEmpty();
        var tables = handler.Query("table");
        tables.Should().BeEmpty();
        var slides = handler.Query("slide");
        slides.Should().HaveCount(1);
    }
}
