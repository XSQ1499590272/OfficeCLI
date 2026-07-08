// Copyright 2026 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Verifies the SDT schema ↔ runtime are honest and aligned:
///   - sdt.json only advertises type values the Add handler actually implements
///   - SDT Add accepts the canonical "type" key (not just the legacy "sdtType")
///   - SDT Get returns the canonical "type" key (not "sdtType")
/// </summary>
[Trait("Speed", "Functional")]
public class WordSdtSchemaHonestyTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"sdthonesty_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var p in _tempFiles)
        {
            try { if (File.Exists(p)) File.Delete(p); } catch { }
        }
    }

    private static string LocateSchema()
    {
        // Walk up from the test assembly dir to find the repo root (has schemas/).
        var dir = AppContext.BaseDirectory;
        for (int i = 0; i < 10 && dir != null; i++)
        {
            var candidate = Path.Combine(dir, "schemas", "help", "docx", "sdt.json");
            if (File.Exists(candidate)) return candidate;
            dir = Path.GetDirectoryName(dir);
        }
        throw new FileNotFoundException("Could not locate schemas/help/docx/sdt.json");
    }

    [Fact]
    public void Schema_TypeValues_OnlyIncludeImplementedVariants()
    {
        var schemaPath = LocateSchema();
        using var doc = JsonDocument.Parse(File.ReadAllText(schemaPath));
        var typeNode = doc.RootElement.GetProperty("properties").GetProperty("type");
        var values = typeNode.GetProperty("values")
            .EnumerateArray().Select(e => e.GetString()).ToList();

        // The implementation supports: text, richtext, dropdown, combobox, date.
        // picture and checkbox are not implemented → must not appear.
        values.Should().NotContain("picture", "picture SDT is not implemented");
        values.Should().NotContain("checkbox", "checkbox SDT is not implemented");

        values.Should().Contain(new[] { "text", "richtext", "dropdown", "combobox", "date" });

        // Set must be false — the builder cannot change variant post-creation.
        typeNode.GetProperty("set").GetBoolean().Should().BeFalse(
            "SDT variant cannot be changed after creation; schema must reflect this");
    }

    [Fact]
    public void Add_AcceptsCanonicalTypeKey_And_GetReturnsCanonicalTypeKey()
    {
        var path = CreateTemp("docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        var sdtPath = handler.Add("/body", "sdt", null, new()
        {
            ["type"] = "dropdown",
            ["alias"] = "Country",
            ["items"] = "US,CN,JP",
            ["text"] = "US"
        });

        sdtPath.Should().NotBeNull();
        var node = handler.Get(sdtPath!);
        node.Should().NotBeNull();
        node!.Type.Should().Be("sdt");

        // Canonical key on Get is "type", not "sdtType".
        node.Format.Should().ContainKey("type");
        node.Format["type"].Should().Be("dropdown");
        node.Format.Should().NotContainKey("sdtType",
            "Get must normalize to the canonical 'type' key (no legacy alias duplication)");
    }

    [Theory]
    [InlineData("unlocked", true)]
    [InlineData("sdtLocked", true)]
    [InlineData("contentLocked", false)]
    [InlineData("sdtContentLocked", false)]
    public void Get_Editable_ReflectsContentLock(string lockVal, bool expectedEditable)
    {
        var path = CreateTemp("docx");
        BlankDocCreator.Create(path);

        string sdtPath;
        using (var handler = new WordHandler(path, editable: true))
        {
            sdtPath = handler.Add("/body", "sdt", null, new()
            {
                ["type"] = "text",
                ["text"] = "x",
                ["lock"] = lockVal
            })!;
            sdtPath.Should().NotBeNull();

            var node = handler.Get(sdtPath);
            node!.Format.Should().ContainKey("editable");
            node.Format["editable"].Should().Be(expectedEditable,
                $"content-locked SDT (lock={lockVal}) editable must reflect content read-only state");
            // lock readback stays correct regardless of editable.
            node.Format["lock"].Should().Be(lockVal);
        }

        // Persistence: reopen the saved file and re-verify.
        using (var reopened = new WordHandler(path, editable: false))
        {
            var node = reopened.Get(sdtPath);
            node!.Format["editable"].Should().Be(expectedEditable,
                "editable derivation must survive reopen");
            node.Format["lock"].Should().Be(lockVal);
        }
    }

    [Fact]
    public void Add_UnsupportedType_Checkbox_Throws()
    {
        var path = CreateTemp("docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        var act = () => handler.Add("/body", "sdt", null, new()
        {
            ["type"] = "checkbox"
        });

        act.Should().Throw<NotSupportedException>()
            .WithMessage("*checkbox*not implemented*");
    }
}
