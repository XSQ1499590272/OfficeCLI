using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordShapeContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddShape_ReadsBackRawDrawingTreeAndGeometry()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var shapePath = handler.Add("/body", "shape", null, new()
        {
            ["geometry"] = "roundRect",
            ["width"] = "4cm",
            ["height"] = "2cm",
            ["fill"] = "EAF2FF",
            ["line"] = "solid;1pt;2B579A",
            ["hPosition"] = "2cm",
            ["vPosition"] = "1cm",
            ["hRelative"] = "page",
            ["vRelative"] = "paragraph",
            ["alt"] = "Decision box"
        });

        var shape = handler.Get(shapePath, depth: 1);
        var geometry = handler.Get($"{shapePath}/spPr[1]/prstGeom[1]");
        var fillColor = handler.Get($"{shapePath}/spPr[1]/solidFill[1]/srgbClr[1]");
        var extents = handler.Get($"{shapePath}/spPr[1]/xfrm[1]/ext[1]");

        Assert.Equal("/body/shape[1]", shapePath);
        Assert.Equal("wsp", shape.Type);
        Assert.Equal(3, shape.ChildCount);
        Assert.Contains(shape.Children, child => child.Path == $"{shapePath}/spPr[1]");
        Assert.Equal("roundRect", Fmt(geometry)["prst"]);
        Assert.Equal("EAF2FF", Fmt(fillColor)["val"]);
        Assert.Equal("1440000", Fmt(extents)["cx"]);
        Assert.Equal("720000", Fmt(extents)["cy"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void SetShape_UpdatesCuratedShapeSurfaceAndReportsAddOnlyKeys()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var shapePath = handler.Add("/body", "shape", null, new()
        {
            ["geometry"] = "rect",
            ["width"] = "4cm",
            ["height"] = "2cm",
            ["fill"] = "EAF2FF"
        });

        var unsupported = handler.Set(shapePath, new()
        {
            ["preset"] = "ellipse",
            ["width"] = "5cm",
            ["height"] = "3cm",
            ["fillcolor"] = "FFECEC",
            ["linecolor"] = "C00000",
            ["linewidth"] = "2pt",
            ["hPosition"] = "3cm"
        });
        var geometry = handler.Get($"{shapePath}/spPr[1]/prstGeom[1]");
        var fillColor = handler.Get($"{shapePath}/spPr[1]/solidFill[1]/srgbClr[1]");
        var extents = handler.Get($"{shapePath}/spPr[1]/xfrm[1]/ext[1]");

        Assert.Contains("hPosition", unsupported);
        Assert.Equal("ellipse", Fmt(geometry)["prst"]);
        Assert.Equal("FFECEC", Fmt(fillColor)["val"]);
        Assert.Equal("1800000", Fmt(extents)["cx"]);
        Assert.Equal("1080000", Fmt(extents)["cy"]);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void RemoveShape_DeletesShapeAndLeavesDocumentValid()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var shapePath = handler.Add("/body", "shape", null, new()
        {
            ["geometry"] = "rect",
            ["width"] = "2cm",
            ["height"] = "1cm"
        });

        handler.Remove(shapePath);

        Assert.Throws<ArgumentException>(() => handler.Get(shapePath));
        Assert.DoesNotContain(handler.Get("/body", depth: 1).Children, child => child.Path.Contains("/shape["));
        Assert.Empty(handler.Validate());
    }
}
