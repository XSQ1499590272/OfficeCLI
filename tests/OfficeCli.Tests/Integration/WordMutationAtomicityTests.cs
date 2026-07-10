using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Integration;

[Trait("Speed", "Integration")]
public class WordMutationAtomicityTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void FailedSetAndRawSet_LeavePackageBytesUnchanged()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
            setup.Add("/body", "paragraph", null, new() { ["text"] = "atomic target" });

        var before = File.ReadAllBytes(path);
        using (var handler = new WordHandler(path, editable: true))
        {
            Assert.Throws<ArgumentException>(() => handler.Set(
                "/body/p[99]", new() { ["align"] = "center" }));
            Assert.Throws<System.Xml.XmlException>(() => handler.RawSet(
                "/document", "//w:body", "append", "<w:p><w:r>"));
        }

        Assert.Equal(before, File.ReadAllBytes(path));
        using var reopened = new WordHandler(path, editable: false);
        Assert.Equal("atomic target", reopened.Query("paragraph").Single().Text);
        Assert.Empty(reopened.Validate());
    }

    [Fact]
    public void MalformedBatchInput_FailsBeforeAnyItemCanMutateDocument()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
            setup.Add("/body", "paragraph", null, new() { ["text"] = "batch atomic target" });

        var before = File.ReadAllBytes(path);
        using (var handler = new WordHandler(path, editable: true))
        {
            var output = BatchExecutor.ExecuteBatch(
                handler,
                "[{\"command\":\"add\",\"parent\":\"/body\",\"unknown\":\"x\"}]",
                json: true);
            Assert.Contains("\"success\": false", output);
            Assert.Contains("error", output);
        }

        Assert.Equal(before, File.ReadAllBytes(path));
        using var reopened = new WordHandler(path, editable: false);
        Assert.Equal("batch atomic target", reopened.Query("paragraph").Single().Text);
        Assert.Empty(reopened.Validate());
    }

    [Fact]
    public void InvalidMoveCopyAndRemove_LeaveBodyAndPackageUnchanged()
    {
        var path = CreateBlankDocx();
        using (var setup = new WordHandler(path, editable: true))
        {
            setup.Add("/body", "paragraph", null, new() { ["text"] = "first" });
            setup.Add("/body", "paragraph", null, new() { ["text"] = "second" });
        }

        var before = File.ReadAllBytes(path);
        using (var handler = new WordHandler(path, editable: true))
        {
            Assert.Throws<ArgumentException>(() => handler.Move(
                "/body/p[1]", "/body/p[99]", null));
            Assert.Throws<ArgumentException>(() => handler.CopyFrom(
                "/body/p[99]", "/body", null));
            Assert.Throws<ArgumentException>(() => handler.Remove("/body/p[99]"));
        }

        Assert.Equal(before, File.ReadAllBytes(path));
        using var reopened = new WordHandler(path, editable: false);
        Assert.Equal(new[] { "first", "second" }, reopened.Query("paragraph").Select(node => node.Text));
        Assert.Empty(reopened.Validate());
    }
}
