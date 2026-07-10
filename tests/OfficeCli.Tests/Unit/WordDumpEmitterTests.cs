using OfficeCli.Handlers;

namespace OfficeCli.Tests.Unit;

[Trait("Speed", "Unit")]
public class WordDumpEmitterTests : WordTestBase
{
    [Fact]
    public void DumpFullDocument_EmitsResourcesBeforeBodyItems()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Set("/", new()
        {
            ["creator"] = "Dump resource test",
            ["title"] = "Resource order"
        });
        handler.Add("/numbering", "abstractNum", null, new()
        {
            ["id"] = "7",
            ["type"] = "multilevel",
            ["level0.format"] = "decimal",
            ["level0.text"] = "%1."
        });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Body text" });

        var (items, warnings) = WordBatchEmitter.EmitWordWithWarnings(handler);
        var bodyIndex = items.FindIndex(item =>
            item.Command == "add"
            && item.Parent == "/body"
            && item.Type == "p");

        Assert.True(bodyIndex > 0);
        Assert.Empty(warnings);
        Assert.True(items.FindIndex(item =>
            item.Command == "raw-set"
            && item.Part == "/numbering") < bodyIndex);
        Assert.True(items.FindIndex(item =>
            item.Command == "raw-set"
            && item.Part == "/styles") < bodyIndex);
        Assert.True(items.FindIndex(item =>
            item.Command == "raw-set"
            && item.Part == "/theme") < bodyIndex);
        Assert.True(items.FindIndex(item =>
            item.Command == "raw-set"
            && item.Part == "/settings") < bodyIndex);
        Assert.True(items.FindIndex(item =>
            item.Command == "raw-set"
            && item.Part == "/docProps/core.xml") < bodyIndex);
        Assert.True(items.FindIndex(item =>
            item.Command == "raw-set"
            && item.Part == "/docProps/app.xml") < bodyIndex);
    }

    [Fact]
    public void DumpFullDocument_EmitsChartOleWithoutAuxiliaryWarnings()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "Q1,Q2",
            ["data"] = "Revenue:10,20",
            ["title"] = "Revenue"
        });
        handler.Add("/body", "ole", null, new()
        {
            ["src"] = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,SGVsbG8=",
            ["oleKind"] = "package",
            ["contentType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ["embedExt"] = "docx",
            ["progId"] = "Word.Document.12",
            ["display"] = "icon",
            ["name"] = "Embedded Word"
        });

        var (items, warnings) = WordBatchEmitter.EmitWordWithWarnings(handler);

        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Type == "chart"
            && item.Props?["data"] == "Revenue:10,20"
            && item.Props?["chartType"] == "column");
        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Type == "ole"
            && item.Props?["progId"] == "Word.Document.12"
            && item.Props?["contentType"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        Assert.Empty(warnings);
    }

    [Fact]
    public void DumpFullDocument_EmitsComplexInlinedPartsCarrierProps()
    {
        var path = CreateBlankDocx();
        var chartXml = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:lang val="en-US"/>
              <c:chart>
                <c:plotArea/>
              </c:chart>
              <c:externalData r:id="rWorkbook"/>
            </c:chartSpace>
            """));
        var chartStyleXml = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"/>
            """));

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(paragraphPath, "inlinedparts", null, new()
        {
            ["runXml"] = """
                <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <w:drawing>
                    <wp:inline>
                      <wp:extent cx="914400" cy="914400"/>
                      <wp:docPr id="1" name="Dumped inlined chart"/>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                          <c:chart r:id="rOld"/>
                        </a:graphicData>
                      </a:graphic>
                    </wp:inline>
                  </w:drawing>
                </w:r>
                """,
            ["part1.relId"] = "rOld",
            ["part1.data"] = $"data:application/vnd.openxmlformats-officedocument.drawingml.chart+xml;base64,{chartXml}",
            ["part1.child1.relId"] = "rStyle",
            ["part1.child1.data"] = $"data:application/vnd.ms-office.chartstyle+xml;base64,{chartStyleXml}",
            ["part1.ext1.relId"] = "rWorkbook",
            ["part1.ext1.type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
            ["part1.ext1.target"] = "embeddings/oleObject1.bin"
        });

        var items = WordBatchEmitter.EmitWord(handler, "/body");
        var carrier = Assert.Single(items, item => item.Command == "add" && item.Type == "inlinedparts");
        var props = carrier.Props!;

        Assert.Equal("/body/p[last()]", carrier.Parent);
        Assert.Contains("<c:chart", props["runXml"]);
        Assert.True(props.ContainsKey("part1.relId"));
        Assert.StartsWith("data:application/vnd.openxmlformats-officedocument.drawingml.chart+xml;base64,", props["part1.data"]);
        Assert.Equal("rStyle", props["part1.child1.relId"]);
        Assert.StartsWith("data:application/vnd.ms-office.chartstyle+xml;base64,", props["part1.child1.data"]);
        Assert.Equal("rWorkbook", props["part1.ext1.relId"]);
        Assert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", props["part1.ext1.type"]);
        Assert.Equal("embeddings/oleObject1.bin", props["part1.ext1.target"]);
    }

    [Fact]
    public void DumpParagraph_EmitsExplicitRunHyperlinkAndFieldItems()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        var paragraphPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        handler.Add(paragraphPath, "run", null, new()
        {
            ["text"] = "Plain",
            ["bold"] = "true"
        });
        handler.Add(paragraphPath, "hyperlink", null, new()
        {
            ["text"] = "Link",
            ["url"] = "https://example.com"
        });
        handler.Add(paragraphPath, "field", null, new()
        {
            ["fieldType"] = "page",
            ["text"] = "9"
        });

        var items = WordBatchEmitter.EmitWord(handler, paragraphPath);

        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Parent == "/body"
            && item.Type == "p");
        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Parent == "/body/p[last()]"
            && item.Type == "r"
            && item.Props?["text"] == "Plain"
            && item.Props?["bold"] == "true");
        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Parent == "/body/p[last()]"
            && item.Type == "hyperlink"
            && item.Props?["text"] == "Link"
            && item.Props?["url"] == "https://example.com");
        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Parent == "/body/p[last()]"
            && item.Type == "field"
            && item.Props?["fieldType"] == "PAGE"
            && item.Props?["text"] == "9");
    }

    [Fact]
    public void DumpTable_EmitsMergedCellProps()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new()
        {
            ["data"] = "A,B,C;D,E,F",
            ["colWidths"] = "1000,1000,1000"
        });
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["fill"] = "F1FAEE",
            ["gridspan"] = "2"
        });
        handler.Set("/body/tbl[1]/tr[2]/tc[1]", new() { ["vmerge"] = "continue" });

        var items = WordBatchEmitter.EmitWord(handler, "/body/tbl[1]");

        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Parent == "/body"
            && item.Type == "table"
            && item.Props?["rows"] == "2"
            && item.Props?["cols"] == "3"
            && item.Props?["colWidths"] == "1000dxa,1000dxa,1000dxa");
        Assert.Contains(items, item =>
            item.Command == "set"
            && item.Path == "/body/tbl[last()]/tr[1]/tc[1]"
            && item.Props?["fill"] == "#F1FAEE"
            && item.Props?["colspan"] == "2"
            && item.Props?["skipGridSync"] == "true");
        Assert.Contains(items, item =>
            item.Command == "set"
            && item.Path == "/body/tbl[last()]/tr[2]/tc[1]"
            && item.Props?["vmerge"] == "continue");
    }

    [Fact]
    public void DumpTextbox_EmitsTextboxAndInnerParagraphText()
    {
        var path = CreateBlankDocx();

        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "textbox", null, new()
        {
            ["text"] = "Sidebar note",
            ["width"] = "4cm",
            ["height"] = "2cm",
            ["fill"] = "EAF2FF",
            ["line.color"] = "2B579A",
            ["line.width"] = "1pt"
        });

        var items = WordBatchEmitter.EmitWord(handler, "/body");

        Assert.Contains(items, item =>
            item.Command == "add"
            && item.Parent == "/body"
            && item.Type == "textbox"
            && item.Props?["width"] == "1440000emu"
            && item.Props?["height"] == "720000emu"
            && item.Props?["fill"] == "EAF2FF"
            && item.Props?["line.color"] == "2B579A");
        Assert.Contains(items, item =>
            item.Command == "set"
            && item.Path == "/body/textbox[1]/p[last()]"
            && item.Props?["text"] == "Sidebar note");
    }
}
