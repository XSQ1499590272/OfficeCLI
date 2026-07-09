using OfficeCli;

namespace OfficeCli.Tests.Unit;

public abstract class WordTestBase : IDisposable
{
    private readonly List<string> _tempFiles = new();

    protected string CreateBlankDocx(string? locale = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_word_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path, locale);
        return path;
    }

    protected static IReadOnlyDictionary<string, object?> Fmt(OfficeCli.Core.DocumentNode node) => node.Format;

    public void Dispose()
    {
        foreach (var path in _tempFiles)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }
}
