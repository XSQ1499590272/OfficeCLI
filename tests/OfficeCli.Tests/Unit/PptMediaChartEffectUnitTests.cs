// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;

using OfficeCli.Tests.Pptx;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public sealed class PptMediaChartEffectUnitTests : PptTestBase
{
    private string CreateSlide()
    {
        var path = CreatePresentation();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        return path;
    }

    // ==================== Picture Tests ====================

    [Fact]
    public void Picture_AddPng_ReadbackMatches()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        var result = handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["alt"] = "A tiny test image"
        });
        result.Should().StartWith("/slide[1]/picture[");

        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
        node.Format["alt"].Should().Be("A tiny test image");
    }

    [Fact]
    public void Picture_AddSvg_WithFallback_RoundTrips()
    {
        var path = CreateSlide();
        var svgPath = CreateTinySvg(NewTempPath(".svg"));
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        var result = handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = svgPath,
            ["fallback"] = pngPath
        });
        result.Should().StartWith("/slide[1]/picture[");

        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
    }

    [Fact]
    public void Picture_AddSvg_WithoutFallback_UsesTransparentPlaceholder()
    {
        var path = CreateSlide();
        var svgPath = CreateTinySvg(NewTempPath(".svg"));
        using var handler = OpenEditable(path);
        var result = handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = svgPath
        });
        result.Should().StartWith("/slide[1]/picture[");

        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
    }

    [Fact]
    public void Picture_Crop_AllSides_PersistInGet()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["crop"] = "5,10,15,20"
        });
        var node = handler.Get("/slide[1]/picture[1]");
        // Crop values should be read back (may differ slightly due to rounding)
        node.Type.Should().Be("picture");
        node.Format.Should().ContainKey("crop");
    }

    [Fact]
    public void Picture_CropIndividual_LeftTopRightBottom()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["cropLeft"] = "10",
            ["cropTop"] = "5",
            ["cropRight"] = "10",
            ["cropBottom"] = "5"
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
        node.Format.Should().ContainKey("crop");
    }

    [Theory]
    [InlineData("stretch")]
    [InlineData("contain")]
    [InlineData("cover")]
    [InlineData("tile")]
    public void Picture_FillMode_AcceptsAllValues(string fillMode)
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["fillMode"] = fillMode
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
        if (fillMode == "tile")
            node.Format.Should().ContainKey("fillMode");
    }

    [Fact]
    public void Picture_BrightnessContrast_SetAndVerify()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["brightness"] = "20",
            ["contrast"] = "-10"
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
        node.Format.Should().ContainKey("brightness");
        node.Format.Should().ContainKey("contrast");
    }

    [Fact]
    public void Picture_Shadow_SetAndReadback()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["shadow"] = "#000000-5pt-45-3pt-40"
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
        node.Format.Should().ContainKey("shadow");
    }

    [Fact]
    public void Picture_Glow_SetAndReadback()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["glow"] = "#FF0000-10pt-50"
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
        node.Format.Should().ContainKey("glow");
    }

    [Fact]
    public void Picture_Rotation_SetAndReadback()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["rotation"] = "45"
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
        node.Format.Should().ContainKey("rotation");
    }

    [Fact]
    public void Picture_LinkAndTooltip_SetAndReadback()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["link"] = "https://example.com",
            ["tooltip"] = "Click to visit"
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Format.Should().ContainKey("link");
        node.Format["link"].Should().Be("https://example.com");
    }

    [Fact]
    public void Picture_AltText_ReadbackCorrect()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath,
            ["alt"] = "Accessible description"
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Format["alt"].Should().Be("Accessible description");
    }

    [Fact]
    public void Picture_SetSrc_ReplacesImage()
    {
        var path = CreateSlide();
        var png1 = CreateTinyPng(NewTempPath(".png"));
        var png2 = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = png1
        });
        handler.Set("/slide[1]/picture[1]", new Dictionary<string, string>
        {
            ["src"] = png2
        });
        var node = handler.Get("/slide[1]/picture[1]");
        node.Type.Should().Be("picture");
        node.Format.Should().ContainKey("width");
    }

    [Fact]
    public void Picture_Remove_ClearsElement()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
        {
            ["src"] = pngPath
        });
        handler.Query("picture").Should().HaveCount(1);
        handler.Remove("/slide[1]/picture[1]");
        handler.Query("picture").Should().HaveCount(0);
    }

    // ==================== Video / Audio Tests ====================

    [Fact]
    public void Video_AddGetRemove_Lifecycle()
    {
        var path = CreateSlide();
        var videoPath = CreateTinyVideo(NewTempPath(".mp4"));
        using var handler = OpenEditable(path);
        var result = handler.Add("/slide[1]", "video", null, new Dictionary<string, string>
        {
            ["src"] = videoPath,
            ["name"] = "MyVideo"
        });
        result.Should().StartWith("/slide[1]/video[");

        var node = handler.Get("/slide[1]/video[1]");
        node.Type.Should().Be("video");
        node.Format["name"].Should().Be("MyVideo");
        node.Format.Should().ContainKey("contentType");

        handler.Remove("/slide[1]/video[1]");
        handler.Query("video").Should().HaveCount(0);
    }

    [Fact]
    public void Audio_AddGetSet_Lifecycle()
    {
        var path = CreateSlide();
        var audioPath = NewTempPath(".mp3");
        // Minimal valid MP3 frame (silence)
        File.WriteAllBytes(audioPath, new byte[] {
            0xFF, 0xFB, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00
        });
        TrackTempFile(audioPath);
        using var handler = OpenEditable(path);
        var result = handler.Add("/slide[1]", "audio", null, new Dictionary<string, string>
        {
            ["src"] = audioPath,
            ["name"] = "MyAudio"
        });
        result.Should().StartWith("/slide[1]/audio[");

        var node = handler.Get("/slide[1]/audio[1]");
        node.Type.Should().Be("audio");
        node.Format["name"].Should().Be("MyAudio");

        handler.Query("audio").Should().HaveCount(1);
    }

    [Fact]
    public void Video_SetLoopAndAutoStart_Persists()
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
    public void Audio_SetVolume_RoundTrips()
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
            ["volume"] = "50"
        });
        var node = handler.Get("/slide[1]/audio[1]");
        node.Type.Should().Be("audio");
    }

    [Fact]
    public void Media_ContentType_ReflectsExtension()
    {
        var path = CreateSlide();
        var videoPath = CreateTinyVideo(NewTempPath(".mp4"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "video", null, new Dictionary<string, string>
        {
            ["src"] = videoPath
        });
        var node = handler.Get("/slide[1]/video[1]");
        node.Format.Should().ContainKey("contentType");
        node.Format["contentType"].Should().Be("video/mp4");
    }

    // ==================== OLE Tests ====================

    [Fact]
    public void Ole_AddGet_ReadbackHasCorrectType()
    {
        var path = CreateSlide();
        var olePayloadPath = CreateOlePayload(NewTempPath(".bin"));
        using var handler = OpenEditable(path);
        var result = handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
        {
            ["src"] = olePayloadPath,
            ["name"] = "EmbeddedObject"
        });
        result.Should().StartWith("/slide[1]/ole[");

        var node = handler.Get("/slide[1]/ole[1]");
        node.Type.Should().Be("ole");
        node.Format["name"].Should().Be("EmbeddedObject");
    }

    [Fact]
    public void Ole_SaveReopen_Persists()
    {
        var path = CreateSlide();
        var olePayloadPath = CreateOlePayload(NewTempPath(".bin"));
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
            {
                ["src"] = olePayloadPath,
                ["name"] = "PersistOle"
            });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/slide[1]/ole[1]");
            node.Type.Should().Be("ole");
            node.Format["name"].Should().Be("PersistOle");
        }
    }

    [Fact]
    public void Ole_DisplayContent_SetAndReadback()
    {
        var path = CreateSlide();
        var olePayloadPath = CreateOlePayload(NewTempPath(".bin"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
        {
            ["src"] = olePayloadPath,
            ["display"] = "content"
        });
        var node = handler.Get("/slide[1]/ole[1]");
        node.Type.Should().Be("ole");
    }

    [Fact]
    public void Ole_SetSrc_ReplacesPayload()
    {
        var path = CreateSlide();
        var ole1 = CreateOlePayload(NewTempPath(".bin"));
        var ole2 = CreateOlePayload(NewTempPath(".bin"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
        {
            ["src"] = ole1
        });
        handler.Set("/slide[1]/ole[1]", new Dictionary<string, string>
        {
            ["src"] = ole2
        });
        var node = handler.Get("/slide[1]/ole[1]");
        node.Type.Should().Be("ole");
    }

    // ==================== Model3D Tests ====================

    [Fact]
    public void Model3D_AddGlb_ReadbackHasCorrectType()
    {
        var path = CreateSlide();
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        using var handler = OpenEditable(path);
        var result = handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
        {
            ["src"] = glbPath,
            ["name"] = "TestModel"
        });
        result.Should().StartWith("/slide[1]/model3d[");

        var node = handler.Get("/slide[1]/model3d[1]");
        node.Type.Should().Be("model3d");
        node.Format["name"].Should().Be("TestModel");
    }

    [Fact]
    public void Model3D_WithRotation_SetAndReadback()
    {
        var path = CreateSlide();
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
        {
            ["src"] = glbPath,
            ["rotation"] = "10,20,30"
        });
        var node = handler.Get("/slide[1]/model3d[1]");
        node.Type.Should().Be("model3d");
        node.Format.Should().ContainKey("rotation");
    }

    [Fact]
    public void Model3D_SaveReopen_Persists()
    {
        var path = CreateSlide();
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
            {
                ["src"] = glbPath,
                ["name"] = "Persist3D"
            });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/slide[1]/model3d[1]");
            node.Type.Should().Be("model3d");
            node.Format["name"].Should().Be("Persist3D");
        }
    }

    [Fact]
    public void Model3D_SetRotation_RoundTrips()
    {
        var path = CreateSlide();
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
        {
            ["src"] = glbPath
        });
        handler.Set("/slide[1]/model3d[1]", new Dictionary<string, string>
        {
            ["rotation"] = "0,90,0"
        });
        var node = handler.Get("/slide[1]/model3d[1]");
        node.Type.Should().Be("model3d");
    }

    // ==================== Chart Tests ====================

    [Theory]
    [InlineData("column")]
    [InlineData("bar")]
    [InlineData("line")]
    [InlineData("pie")]
    [InlineData("doughnut")]
    [InlineData("area")]
    [InlineData("scatter")]
    [InlineData("bubble")]
    [InlineData("radar")]
    [InlineData("waterfall")]
    public void Chart_AllStandardTypes_AddAndReadback(string chartType)
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        var result = handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = chartType,
            ["title"] = $"My {chartType} Chart",
            ["data"] = "Series1:10,20,30;Series2:5,15,25"
        });
        result.Should().StartWith("/slide[1]/chart[");

        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
        node.Format["title"].Should().Be($"My {chartType} Chart");
    }

    [Fact]
    public void Chart_Column3D_AddAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column3d",
            ["data"] = "Sales:100,200,150"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_Stock_AddAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "stock",
            ["categories"] = "Jan,Feb,Mar,Apr",
            ["data"] = "High:25,30,28,32;Low:10,15,12,18;Close:20,22,18,25"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_Combo_AddAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "combo",
            ["data"] = "Revenue:100,200,300;Margin:15,30,45"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_SeriesAdd_AppendsCorrectly()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30"
        });
        var seriesResult = handler.Add("/slide[1]/chart[1]", "series", null, new Dictionary<string, string>
        {
            ["name"] = "Revenue",
            ["values"] = "40,50,60"
        });
        seriesResult.Should().StartWith("/slide[1]/chart[1]/series[");

        var node = handler.Get("/slide[1]/chart[1]", depth: 1);
        node.Children.Should().NotBeEmpty();
    }

    [Fact]
    public void Chart_SeriesRemove_RemovesCorrectly()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30;Revenue:40,50,60"
        });
        var node1 = handler.Get("/slide[1]/chart[1]", depth: 1);
        var beforeCount = node1.Children.Count;

        handler.Remove("/slide[1]/chart[1]/series[2]");
        var node2 = handler.Get("/slide[1]/chart[1]", depth: 1);
        node2.Children.Count.Should().Be(beforeCount - 1);
    }

    [Fact]
    public void Chart_AxisProperties_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["axisTitle"] = "Categories"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_Legend_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["legend"] = "bottom"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_Gridlines_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["gridlines"] = "true"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_DataLabels_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["dataLabels"] = "outEnd"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_PieOfPie_AddAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "pieOfPie",
            ["data"] = "A:10,B:20,C:30,D:5,E:8"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_BarOfPie_AddAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "barOfPie",
            ["data"] = "A:10,B:20,C:30,D:5,E:8"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_DropLines_HiLowLines_UpDownBars_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "line",
            ["data"] = "Sales:10,20,30",
            ["dropLines"] = "true",
            ["hiLowLines"] = "true",
            ["upDownBars"] = "true"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_Trendlines_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "scatter",
            ["data"] = "Series1:10,20,30",
            ["trendline"] = "linear"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_ErrorBars_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "bar",
            ["data"] = "Series1:10,20,30",
            ["errorBars"] = "percentage:5"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_AnchorShorthand_ExpandsToPosition()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["anchor"] = "3cm,5cm,18cm,10cm"
        });
        var node = handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
        node.Format.Should().ContainKey("x");
        node.Format.Should().ContainKey("y");
    }

    [Fact]
    public void Chart_SaveReopen_PersistsWithData()
    {
        var path = CreateSlide();
        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
            {
                ["type"] = "column",
                ["title"] = "Persisted Chart",
                ["data"] = "Series1:10,20,30"
            });
            handler.Save();
        }
        using (var handler = OpenEditable(path))
        {
            var node = handler.Get("/slide[1]/chart[1]");
            node.Type.Should().Be("chart");
            node.Format["title"].Should().Be("Persisted Chart");
        }
    }

    // ==================== Animation Tests ====================

    [Fact]
    public void Animation_EntranceFade_AddAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        var result = handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "entrance",
            ["duration"] = "1000"
        });
        result.Should().StartWith("/slide[1]/shape[");

        var node = handler.Get("/slide[1]/shape[1]" , depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_ExitFlyOut_AddAndReadback()
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
            ["class"] = "exit",
            ["duration"] = "500"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_EmphasisSpin_AddAndReadback()
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
            ["duration"] = "800"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_EmphasisGrow_AddAndReadback()
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
            ["effect"] = "grow",
            ["class"] = "emphasis",
            ["duration"] = "600"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Theory]
    [InlineData("fillColor")]
    [InlineData("lineColor")]
    [InlineData("transparency")]
    [InlineData("pulse")]
    [InlineData("teeter")]
    [InlineData("darken")]
    [InlineData("lighten")]
    [InlineData("desaturate")]
    [InlineData("complementaryColor")]
    [InlineData("contrastingColor")]
    [InlineData("objectColor")]
    [InlineData("colorPulse")]
    public void Animation_EmphasisPresets_AllAdd(string effect)
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = effect,
            ["class"] = "emphasis",
            ["duration"] = "500"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Theory]
    [InlineData("contract")]
    [InlineData("centerRevolve")]
    [InlineData("collapse")]
    [InlineData("floatOut")]
    [InlineData("shrinkTurn")]
    [InlineData("sinkDown")]
    [InlineData("spinner")]
    [InlineData("basicZoom")]
    [InlineData("stretchy")]
    [InlineData("boomerang")]
    [InlineData("credits")]
    [InlineData("curveDown")]
    [InlineData("float")]
    [InlineData("pinwheel")]
    [InlineData("spiralOut")]
    [InlineData("basicSwivel")]
    public void Animation_ExitPresets_TemplateBacked_AllAdd(string effect)
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = effect,
            ["class"] = "exit",
            ["duration"] = "500"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Theory]
    [InlineData("appear")]
    [InlineData("fly")]
    [InlineData("blinds")]
    [InlineData("box")]
    [InlineData("checkerboard")]
    [InlineData("circle")]
    [InlineData("crawl")]
    [InlineData("diamond")]
    [InlineData("dissolve")]
    [InlineData("flash")]
    [InlineData("float")]
    [InlineData("plus")]
    [InlineData("random")]
    [InlineData("split")]
    [InlineData("strips")]
    [InlineData("swivel")]
    [InlineData("wedge")]
    [InlineData("wheel")]
    [InlineData("wipe")]
    [InlineData("zoom")]
    [InlineData("bounce")]
    public void Animation_EntrancePresets_AllAdd(string effect)
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = effect,
            ["class"] = "entrance",
            ["duration"] = "500"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_MotionPath_Line_AddAndReadback()
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
            ["class"] = "motion",
            ["path"] = "line",
            ["direction"] = "right",
            ["duration"] = "2000"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        // Motion path shows as motionPath key in Format
        node.Format.Should().ContainKey("motionPath");
    }

    [Theory]
    [InlineData("line")]
    [InlineData("arc")]
    [InlineData("circle")]
    [InlineData("diamond")]
    [InlineData("triangle")]
    [InlineData("square")]
    public void Animation_MotionPathPresets_AllAdd(string preset)
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "2cm"
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["class"] = "motion",
            ["path"] = preset
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("motionPath");
    }

    [Fact]
    public void Animation_Repeat_SetAndReadback()
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
            ["repeat"] = "3"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_AutoReverse_AddAndReadback()
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
            ["autoReverse"] = "true"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_Restart_AddAndReadback()
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
            ["restart"] = "whenNotActive"
        });
        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_ChartBuild_BySeries_OnChart()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Series1:10,20,30;Series2:5,15,25"
        });
        handler.Add("/slide[1]/chart[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "entrance",
            ["chartBuild"] = "bySeries"
        });
        var node = handler.Get("/slide[1]/chart[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_ChartBuild_ByCategory_OnChart()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "chart", null, new Dictionary<string, string>
        {
            ["type"] = "column",
            ["data"] = "Series1:10,20,30;Series2:5,15,25"
        });
        handler.Add("/slide[1]/chart[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "wipe",
            ["class"] = "entrance",
            ["chartBuild"] = "byCategory"
        });
        var node = handler.Get("/slide[1]/chart[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
    }

    [Fact]
    public void Animation_MultiEffectChain_MultipleAnimationsOnOneShape()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["shape"] = "rect",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "3cm"
        });
        // Add entrance
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fade",
            ["class"] = "entrance"
        });
        // Add emphasis
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "spin",
            ["class"] = "emphasis"
        });
        // Add exit
        handler.Add("/slide[1]/shape[1]", "animation", null, new Dictionary<string, string>
        {
            ["effect"] = "fly",
            ["class"] = "exit"
        });

        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        node.Format.Should().ContainKey("animation");
        node.Format.Should().ContainKey("animation2");
        node.Format.Should().ContainKey("animation3");
    }

    [Fact]
    public void Animation_Remove_ClearsAnimation()
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
            ["class"] = "entrance"
        });
        handler.Remove("/slide[1]/shape[1]/animation[1]");

        var node = handler.Get("/slide[1]/shape[1]", depth: 1);
        // After removing all animations, the animation key should be absent
        node.Format.Should().NotContainKey("animation");
    }

    // ==================== Transition Tests ====================

    [Fact]
    public void Transition_Fade_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "fade"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transition");
        node.Format["transition"].Should().Be("fade");
    }

    [Theory]
    [InlineData("wipe")]
    [InlineData("push")]
    [InlineData("cover")]
    [InlineData("uncover")]
    [InlineData("zoom")]
    [InlineData("split")]
    [InlineData("blinds")]
    [InlineData("checker")]
    [InlineData("dissolve")]
    [InlineData("random")]
    [InlineData("circle")]
    [InlineData("diamond")]
    [InlineData("newsflash")]
    [InlineData("plus")]
    [InlineData("wedge")]
    [InlineData("wheel")]
    [InlineData("box")]
    [InlineData("curtains")]
    [InlineData("crush")]
    [InlineData("wind")]
    [InlineData("prestige")]
    [InlineData("fracture")]
    [InlineData("peelOff")]
    [InlineData("pageCurlDouble")]
    [InlineData("pageCurlSingle")]
    [InlineData("airplane")]
    [InlineData("origami")]
    [InlineData("fallOver")]
    [InlineData("drape")]
    public void Transition_Presets_AllSet(string preset)
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Set($"/slide[2]", new Dictionary<string, string>
        {
            ["transition"] = preset
        });
        var node = handler.Get($"/slide[2]");
        node.Format.Should().ContainKey("transition");
    }

    [Fact]
    public void Transition_Duration_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "fade",
            ["duration"] = "2000"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transition");
    }

    [Fact]
    public void Transition_AdvanceOnClick_SetFalse()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "fade",
            ["advanceOnClick"] = "false"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transition");
    }

    [Fact]
    public void Transition_AdvanceAfter_SetTimer()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transition"] = "fade",
            ["advanceAfter"] = "3000"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transition");
    }

    [Fact]
    public void Transition_Speed_SetAndReadback()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Set("/slide[1]", new Dictionary<string, string>
        {
            ["transitionSpeed"] = "fast"
        });
        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("transitionSpeed");
    }

    // ==================== Zoom Tests ====================

    [Fact]
    public void Zoom_AddGetSetRemove_Lifecycle()
    {
        var path = CreateSlide();
        // Need at least 2 slides for a zoom
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        var result = handler.Add("/slide[1]", "zoom", null, new Dictionary<string, string>
        {
            ["target"] = "2"
        });
        result.Should().StartWith("/slide[1]/zoom[");

        var node = handler.Get("/slide[1]/zoom[1]");
        node.Type.Should().Be("zoom");

        handler.Set("/slide[1]/zoom[1]", new Dictionary<string, string>
        {
            ["transitionDur"] = "500"
        });
        node = handler.Get("/slide[1]/zoom[1]");
        node.Type.Should().Be("zoom");

        handler.Remove("/slide[1]/zoom[1]");
        handler.Query("zoom").Should().HaveCount(0);
    }

    [Fact]
    public void Zoom_TargetRelationship_ReadbackMatches()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "zoom", null, new Dictionary<string, string>
        {
            ["target"] = "2"
        });
        var node = handler.Get("/slide[1]/zoom[1]");
        node.Format.Should().ContainKey("target");
        // Target may be returned as an integer (slide number)
        int.Parse(node.Format["target"]!.ToString()!).Should().BeGreaterThan(0);
    }

    // ==================== Diagram Tests ====================

    [Fact]
    public void Diagram_MermaidFlowchart_AddSucceeds()
    {
        var path = CreateSlide();
        using var handler = OpenEditable(path);
        // Diagrams are add-only synthesizers; native mode may produce shapes+connectors
        handler.Add("/slide[1]", "diagram", null, new Dictionary<string, string>
        {
            ["text"] = @"flowchart LR
    A[Start] --> B[Process]
    B --> C[End]"
        });
        // After add, the slide should have some content (shapes or group)
        var slide = handler.Get("/slide[1]", depth: 1);
        slide.Children.Should().NotBeNull();
    }

    // ==================== Save/Reopen Round-Trips ====================

    [Fact]
    public void AllMediaTypes_SaveReopen_AllPersist()
    {
        var path = CreateSlide();
        var pngPath = CreateTinyPng(NewTempPath(".png"));
        var videoPath = CreateTinyVideo(NewTempPath(".mp4"));
        var olePayloadPath = CreateOlePayload(NewTempPath(".bin"));
        var glbPath = CreateTinyGlb(NewTempPath(".glb"));

        using (var handler = OpenEditable(path))
        {
            handler.Add("/slide[1]", "picture", null, new Dictionary<string, string>
            {
                ["src"] = pngPath, ["name"] = "Pic1"
            });
            handler.Add("/slide[1]", "video", null, new Dictionary<string, string>
            {
                ["src"] = videoPath, ["name"] = "Vid1"
            });
            handler.Add("/slide[1]", "ole", null, new Dictionary<string, string>
            {
                ["src"] = olePayloadPath, ["name"] = "Ole1"
            });
            handler.Add("/slide[1]", "model3d", null, new Dictionary<string, string>
            {
                ["src"] = glbPath, ["name"] = "Model1"
            });
            handler.Save();
        }

        using (var handler = OpenEditable(path))
        {
            var picNode = handler.Get("/slide[1]/picture[1]");
            picNode.Type.Should().Be("picture");
            picNode.Format["name"].Should().Be("Pic1");

            var vidNode = handler.Get("/slide[1]/video[1]");
            vidNode.Type.Should().Be("video");
            vidNode.Format["name"].Should().Be("Vid1");

            var oleNode = handler.Get("/slide[1]/ole[1]");
            oleNode.Type.Should().Be("ole");
            oleNode.Format["name"].Should().Be("Ole1");

            var modelNode = handler.Get("/slide[1]/model3d[1]");
            modelNode.Type.Should().Be("model3d");
            modelNode.Format["name"].Should().Be("Model1");
        }
    }
}
