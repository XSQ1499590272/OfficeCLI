using OfficeCli.Handlers;

namespace OfficeCli.Tests.Functional;

[Trait("Speed", "Functional")]
public class WordOleContractTests : OfficeCli.Tests.Unit.WordTestBase
{
    [Fact]
    public void AddOle_FromDataUriPackage_ReadsBackPayloadMetadataAndFrame()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var olePath = handler.Add("/body", "ole", null, new()
        {
            ["src"] = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,SGVsbG8=",
            ["oleKind"] = "package",
            ["contentType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ["embedExt"] = "docx",
            ["progId"] = "Word.Document.12",
            ["display"] = "icon",
            ["name"] = "Embedded Word",
            ["width"] = "2cm",
            ["height"] = "1cm"
        });

        var ole = handler.Get(olePath);
        var queriedOle = Assert.Single(handler.Query("ole"));
        var fmt = Fmt(ole);

        Assert.Equal(olePath, queriedOle.Path);
        Assert.Equal("ole", ole.Type);
        Assert.Equal("ole", fmt["objectType"]);
        Assert.Equal("Word.Document.12", fmt["progId"]);
        Assert.Equal("icon", fmt["display"]);
        Assert.Equal("Embedded Word", fmt["name"]);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", fmt["contentType"]);
        Assert.Equal(5L, fmt["fileSize"]);
        Assert.Equal("2cm", fmt["width"]);
        Assert.Equal("1cm", fmt["height"]);
        Assert.False(string.IsNullOrWhiteSpace(fmt["relId"]?.ToString()));
        Assert.Empty(handler.Validate());
    }
}
