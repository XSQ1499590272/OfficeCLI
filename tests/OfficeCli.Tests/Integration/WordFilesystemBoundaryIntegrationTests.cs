using OfficeCli;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Collection("Word CLI output")]
[Trait("Speed", "Integration")]
public sealed class WordFilesystemBoundaryIntegrationTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void CreateHandlesUnicodeSpacePath_TypeInference_AndForceOverwrite()
    {
        var previous = Environment.GetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT");
        var directory = Path.Combine(Path.GetTempPath(), $"officecli word 边界 {Guid.NewGuid():N}");
        var extensionless = Path.Combine(directory, "文档 with spaces");
        var document = extensionless + ".docx";

        Environment.SetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT", "1");
        try
        {
            Directory.CreateDirectory(directory);

            Assert.Equal(0, Invoke("create", extensionless, "--type", "docx", "--minimal"));
            Assert.True(File.Exists(document));
            using (var created = new WordHandler(document, editable: false))
                Assert.Empty(created.Validate());

            var before = File.ReadAllBytes(document);
            Assert.NotEqual(0, Invoke("create", document, "--type", "docx", "--minimal"));
            Assert.Equal(before, File.ReadAllBytes(document));

            Assert.Equal(0, Invoke("create", document, "--type", "docx", "--minimal", "--force"));
            using var overwritten = new WordHandler(document, editable: false);
            Assert.Empty(overwritten.Validate());
        }
        finally
        {
            Environment.SetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT", previous);
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); } catch { }
        }
    }

    [Fact]
    public void DumpOutToMissingParentFailsWithoutCreatingPartialOutput()
    {
        var path = CreateBlankDocx();
        var directory = Path.Combine(Path.GetTempPath(), $"officecli_missing_parent_{Guid.NewGuid():N}");
        var output = Path.Combine(directory, "nested", "dump.json");

        try
        {
            var exitCode = Invoke("dump", path, "/", "--out", output);

            Assert.NotEqual(0, exitCode);
            Assert.False(File.Exists(output));
            Assert.False(Directory.Exists(Path.GetDirectoryName(output)));
        }
        finally
        {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); } catch { }
        }
    }

    private static int Invoke(params string[] args)
    {
        var root = CommandBuilder.BuildRootCommand();
        return root.Parse(args).Invoke();
    }
}
