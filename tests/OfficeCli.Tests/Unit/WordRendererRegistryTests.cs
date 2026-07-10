using OfficeCli.Core;
using OfficeCli.Core.Rendering;
using OfficeCli.Handlers.Rendering;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordRendererRegistryTests
{
    [Fact]
    public void RenderCapabilities_CoversFormatOutputAndWatchMode()
    {
        var capabilities = new RenderCapabilities
        {
            Name = "word-test",
            Priority = 10,
            SupportedFormats = ["docx"],
            SupportedOutputs = RenderOutputKind.Html | RenderOutputKind.Svg,
            SupportsWatch = true
        };

        Assert.True(capabilities.Covers("DOCX", RenderOutputKind.Html, RenderMode.Static));
        Assert.True(capabilities.Covers("docx", RenderOutputKind.Svg, RenderMode.Watch));
        Assert.False(capabilities.Covers("xlsx", RenderOutputKind.Html, RenderMode.Static));
        Assert.False(capabilities.Covers("docx", RenderOutputKind.Png, RenderMode.Static));
    }

    [Fact]
    public void RenderCapabilities_RequiresWatchSupportOnlyForWatchRequests()
    {
        var capabilities = new RenderCapabilities
        {
            Name = "static-only",
            Priority = 0,
            SupportedFormats = ["docx"],
            SupportedOutputs = RenderOutputKind.Html,
            SupportsWatch = false
        };

        Assert.True(capabilities.Covers("docx", RenderOutputKind.Html, RenderMode.Static));
        Assert.False(capabilities.Covers("docx", RenderOutputKind.Html, RenderMode.Watch));
    }

    [Fact]
    public void Registry_SelectsHighestPriorityAvailableRenderer()
    {
        var registry = new RendererRegistry();
        var fallback = new TestRenderer("fallback", priority: 0, available: true);
        var preferred = new TestRenderer("preferred", priority: 20, available: true);
        registry.Register(fallback);
        registry.Register(preferred);

        Assert.Same(preferred, registry.Resolve("docx", RenderOutputKind.Html, RenderMode.Static));

        preferred.Available = false;
        Assert.Same(fallback, registry.Resolve("docx", RenderOutputKind.Html, RenderMode.Static));
    }

    [Fact]
    public void Registry_RegisterIsIdempotentPerRendererInstance()
    {
        var registry = new RendererRegistry();
        var renderer = new TestRenderer("same", priority: 0, available: true);

        registry.Register(renderer);
        registry.Register(renderer);

        Assert.Single(registry.Registered());
        Assert.Throws<ArgumentNullException>(() => registry.Register(null!));
    }

    [Fact]
    public void RenderingBootstrap_RegistersWordRendererOnceWithWatchCapability()
    {
        var registry = new RendererRegistry();

        RenderingBootstrap.EnsureRegistered(registry);
        RenderingBootstrap.EnsureRegistered(registry);

        var word = Assert.Single(registry.Registered(),
            capabilities => capabilities.Name == "basic-html-docx");
        Assert.Equal(0, word.Priority);
        Assert.Equal(RenderOutputKind.Html, word.SupportedOutputs);
        Assert.True(word.SupportsWatch);
        Assert.True(registry.Resolve("DOCX", RenderOutputKind.Html, RenderMode.Watch) != null);
    }

    [Fact]
    public void WordBasicRenderer_RejectsNonWordRenderInput()
    {
        var renderer = new WordBasicRenderer();

        var exception = Assert.Throws<InvalidOperationException>(() =>
            renderer.Render(new EmptyRenderInput(), new RenderOptions()));

        Assert.Contains("WordHandler render input expected", exception.Message);
    }

    private sealed class TestRenderer : IRenderer
    {
        public TestRenderer(string name, int priority, bool available)
        {
            Available = available;
            Capabilities = new RenderCapabilities
            {
                Name = name,
                Priority = priority,
                SupportedFormats = ["docx"],
                SupportedOutputs = RenderOutputKind.Html
            };
        }

        public bool Available { get; set; }
        public bool IsAvailable => Available;
        public RenderCapabilities Capabilities { get; }
        public RenderResult Render(IRenderInput input, RenderOptions options) => RenderResult.Html("test");
    }

    private sealed class EmptyRenderInput : IRenderInput
    {
        public string FormatId => "docx";
        public object? Model => null;
        public DocumentNode Get(string path, int depth = 1) => throw new NotSupportedException();
        public IReadOnlyList<DocumentNode> Query(string selector) => throw new NotSupportedException();
    }
}
