using System.Text.Json;
using OfficeCli.Help;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordSchemaParityTests
{
    private static readonly string[] Operations = ["add", "set", "get", "query", "remove"];

    [Fact]
    public void EveryEmbeddedDocxSchemaDeclaresStableOperationAndPropertyShapes()
    {
        var elements = SchemaHelpLoader.ListElements("docx");

        Assert.NotEmpty(elements);
        Assert.Equal(elements.Count, elements.Distinct(StringComparer.OrdinalIgnoreCase).Count());

        foreach (var element in elements)
        {
            using var schema = SchemaHelpLoader.LoadSchema("docx", element);
            var root = schema.RootElement;

            Assert.Equal(JsonValueKind.Object, root.ValueKind);
            Assert.Equal("docx", root.GetProperty("format").GetString());
            Assert.False(string.IsNullOrWhiteSpace(root.GetProperty("element").GetString()),
                $"{element} is missing its canonical element name");

            var operations = root.GetProperty("operations");
            Assert.Equal(JsonValueKind.Object, operations.ValueKind);
            foreach (var operation in operations.EnumerateObject())
            {
                if (string.Equals(operation.Name, "note", StringComparison.OrdinalIgnoreCase))
                {
                    Assert.Equal(JsonValueKind.String, operation.Value.ValueKind);
                    continue;
                }

                Assert.Contains(operation.Name, Operations, StringComparer.OrdinalIgnoreCase);
                Assert.True(operation.Value.ValueKind is JsonValueKind.True or JsonValueKind.False,
                    $"{element}.operations.{operation.Name} must be boolean");
            }

            if (!root.TryGetProperty("properties", out var properties))
                continue;

            Assert.Equal(JsonValueKind.Object, properties.ValueKind);
            foreach (var property in properties.EnumerateObject())
            {
                Assert.True(property.Value.ValueKind == JsonValueKind.Object,
                    $"{element}.{property.Name} must be an object");
                Assert.True(property.Value.TryGetProperty("type", out var type),
                    $"{element}.{property.Name} is missing type");
                Assert.True(type.ValueKind == JsonValueKind.String,
                    $"{element}.{property.Name}.type must be a string");

                foreach (var operation in Operations)
                {
                    if (property.Value.TryGetProperty(operation, out var flag))
                    {
                        Assert.True(flag.ValueKind is JsonValueKind.True or JsonValueKind.False,
                            $"{element}.{property.Name}.{operation} must be boolean");
                    }
                }
            }
        }
    }

    [Theory]
    [InlineData("word", "/", "document")]
    [InlineData("docx", "p", "paragraph")]
    [InlineData("docx", "tbl", "table")]
    [InlineData("docx", "tc", "cell")]
    public void SchemaLoader_ResolvesWordFormatRootAndPathAliases(
        string format, string element, string canonicalElement)
    {
        using var schema = SchemaHelpLoader.LoadSchema(format, element);

        Assert.Equal(canonicalElement, schema.RootElement.GetProperty("element").GetString());
        Assert.Equal("docx", schema.RootElement.GetProperty("format").GetString());
    }
}
