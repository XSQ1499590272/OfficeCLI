// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using FluentAssertions;

namespace OfficeCli.Tests.Excel;

public sealed class ExcelCliCommandE2ETests : ExcelTestBase
{
    // ==================== Core lifecycle ====================

    [Fact]
    public void Create_OpensAndValidates()
    {
        var path = NewTempWorkbookPath();
        var result = RunCliOk("create", path);
        result.Stdout.Should().Contain("Created");
        File.Exists(path).Should().BeTrue();

        var validate = RunCliOk("validate", path);
        validate.Stdout.Should().Contain("no errors");
    }

    [Fact]
    public void Lifecycle_CreateOpenCloseSaveValidate()
    {
        var path = NewTempWorkbookPath();
        RunCliOk("create", path);
        RunCliOk("open", path);
        RunCliOk("save", path);
        RunCliOk("close", path);
        RunCliOk("validate", path);
    }

    [Fact]
    public void Create_WithForce_OverwritesExisting()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("create", path, "--force");
        result.Stderr.Should().Contain("Overwriting");
        File.Exists(path).Should().BeTrue();
        RunCliOk("validate", path);
    }

    [Fact]
    public void Create_WithoutForce_OnExistingFile_Errors()
    {
        var path = CreateWorkbook();
        var result = RunCli("create", path);
        result.ExitCode.Should().NotBe(0);
        (result.Stdout + result.Stderr).Should().Contain("already exists");
    }

    // ==================== add ====================

    [Fact]
    public void Add_Sheet()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=Data");
        result.Stdout.Should().Contain("Added sheet").And.Contain("Data");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/Data");
        getResult.Stdout.Should().Contain("sheet");
    }

    [Fact]
    public void Add_Cell()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=Hello");
        result.Stdout.Should().Contain("Added cell");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/Sheet1/A1");
        getResult.Stdout.Should().Contain("Hello");
    }

    [Fact]
    public void Add_CellWithFormula()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=10", "--prop", "type=number");
        RunCliOk("add", path, "/Sheet1/A2", "--type", "cell", "--prop", "value=20", "--prop", "type=number");
        var result = RunCliOk("add", path, "/Sheet1/A3", "--type", "cell", "--prop", "formula=SUM(A1:A2)");
        result.Stdout.Should().Contain("Added cell");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Add_Row()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("add", path, "/Sheet1", "--type", "row", "--prop", "cols=3");
        result.Stdout.Should().Contain("Added row");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Add_RowWithCellValues()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("add", path, "/Sheet1", "--type", "row", "--prop", "cols=2", "--prop", "c1=Alpha", "--prop", "c2=Beta");
        result.Stdout.Should().Contain("Added row");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/Sheet1/A1");
        getResult.Stdout.Should().Contain("Alpha");
    }

    [Fact]
    public void Add_Table()
    {
        var path = CreateWorkbook();
        // Add some data cells first
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=Col1");
        RunCliOk("add", path, "/Sheet1/B1", "--type", "cell", "--prop", "value=Col2");
        RunCliOk("add", path, "/Sheet1/A2", "--type", "cell", "--prop", "value=1", "--prop", "type=number");
        RunCliOk("add", path, "/Sheet1/B2", "--type", "cell", "--prop", "value=2", "--prop", "type=number");

        var result = RunCliOk("add", path, "/Sheet1", "--type", "table",
            "--prop", "ref=A1:B2", "--prop", "name=MyTable", "--prop", "columns=Col1,Col2");
        result.Stdout.Should().Contain("Added table");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Add_Chart()
    {
        var path = CreateWorkbook();
        // Add data
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=10", "--prop", "type=number");
        RunCliOk("add", path, "/Sheet1/B1", "--type", "cell", "--prop", "value=20", "--prop", "type=number");
        RunCliOk("add", path, "/Sheet1/A2", "--type", "cell", "--prop", "value=30", "--prop", "type=number");
        RunCliOk("add", path, "/Sheet1/B2", "--type", "cell", "--prop", "value=40", "--prop", "type=number");

        var result = RunCliOk("add", path, "/Sheet1", "--type", "chart",
            "--prop", "chartType=column",
            "--prop", "dataRange=Sheet1!A1:B2",
            "--prop", "anchor=F2:L16",
            "--prop", "title=TestChart");
        result.Stdout.Should().Contain("Added chart");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Add_NamedRange()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("add", path, "/", "--type", "namedrange",
            "--prop", "name=MyRange", "--prop", "refersTo=Sheet1!$A$1:$B$2");
        result.Stdout.Should().Contain("Added namedrange");
        RunCliOk("validate", path);
    }

    // ==================== get ====================

    [Fact]
    public void Get_Root_TextOutput()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("get", path, "/");
        result.Stdout.Should().Contain("workbook");
    }

    [Fact]
    public void Get_Sheet()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("get", path, "/Sheet1");
        result.Stdout.Should().Contain("sheet").And.Contain("Sheet1");
    }

    [Fact]
    public void Get_JsonOutput()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("get", path, "/", "--json");
        result.Stdout.Should().Contain("\"type\"");
    }

    [Fact]
    public void Get_Cell()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=TestValue");
        var result = RunCliOk("get", path, "/Sheet1/A1");
        result.Stdout.Should().Contain("TestValue");
    }

    // ==================== query ====================

    [Fact]
    public void Query_Sheet()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("query", path, "sheet");
        result.Stdout.Should().Contain("Sheet1");
    }

    [Fact]
    public void Query_Cell()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=QueryTarget");
        var result = RunCliOk("query", path, "cell");
        result.Stdout.Should().Contain("QueryTarget");
    }

    [Fact]
    public void Query_Table()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=H1");
        RunCliOk("add", path, "/Sheet1/B1", "--type", "cell", "--prop", "value=H2");
        RunCliOk("add", path, "/Sheet1/A2", "--type", "cell", "--prop", "value=1", "--prop", "type=number");
        RunCliOk("add", path, "/Sheet1", "--type", "table", "--prop", "ref=A1:B2", "--prop", "name=QTbl", "--prop", "columns=H1,H2");

        var result = RunCliOk("query", path, "table");
        result.Stdout.Should().Contain("QTbl");
    }

    [Fact]
    public void Query_Sheet_Json()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("query", path, "sheet", "--json");
        result.Stdout.Should().Contain("\"Sheet1\"");
    }

    // ==================== set ====================

    [Fact]
    public void Set_CellValue()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=Original");
        var result = RunCliOk("set", path, "/Sheet1/A1", "--prop", "value=Updated");
        result.Stdout.Should().Contain("Updated");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/Sheet1/A1");
        getResult.Stdout.Should().Contain("Updated");
    }

    [Fact]
    public void Set_CellBold()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=BoldText");
        var result = RunCliOk("set", path, "/Sheet1/A1", "--prop", "bold=true");
        result.Stdout.Should().Contain("Updated");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Set_SheetProperty()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("set", path, "/Sheet1", "--prop", "name=RenamedSheet");
        result.Stdout.Should().Contain("Updated");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/RenamedSheet");
        getResult.Stdout.Should().Contain("sheet");
    }

    // ==================== remove ====================

    [Fact]
    public void Remove_Cell()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=ToDelete");
        var result = RunCliOk("remove", path, "/Sheet1/A1");
        result.Stdout.Should().Contain("Removed");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Remove_Sheet()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=ToRemove");
        var result = RunCliOk("remove", path, "/ToRemove");
        result.Stdout.Should().Contain("Removed");
        RunCliOk("validate", path);
    }

    // ==================== move ====================

    [Fact]
    public void Move_SheetReorder()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=First");
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=Second");

        var result = RunCliOk("move", path, "/Second", "--index", "0");
        result.Stdout.Should().Contain("Moved");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Move_Row()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1", "--type", "row", "--prop", "cols=1", "--prop", "c1=R1");
        RunCliOk("add", path, "/Sheet1", "--type", "row", "--prop", "cols=1", "--prop", "c1=R2");
        RunCliOk("add", path, "/Sheet1", "--type", "row", "--prop", "cols=1", "--prop", "c1=R3");

        // Move row[2] after row[3]
        var result = RunCliOk("move", path, "/Sheet1/row[2]", "--to", "/Sheet1", "--after", "/Sheet1/row[3]");
        result.Stdout.Should().Contain("Moved");
        RunCliOk("validate", path);
    }

    // ==================== swap ====================

    [Fact]
    public void Swap_Rows()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1", "--type", "row", "--prop", "cols=1", "--prop", "c1=First");
        RunCliOk("add", path, "/Sheet1", "--type", "row", "--prop", "cols=1", "--prop", "c1=Second");

        var result = RunCliOk("swap", path, "/Sheet1/row[1]", "/Sheet1/row[2]");
        result.Stdout.Should().Contain("Swapped");
        RunCliOk("validate", path);

        var getA1 = RunCliOk("get", path, "/Sheet1/A1");
        getA1.Stdout.Should().Contain("Second");
    }

    // ==================== batch ====================

    [Fact]
    public void Batch_CommandsJson()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("batch", path, "--commands", """
            [
              {"command":"add","parent":"/Sheet1/A1","type":"cell","props":{"value":"BatchCell"}},
              {"command":"set","path":"/Sheet1/A1","props":{"bold":"true"}}
            ]
            """, "--json");
        result.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/Sheet1/A1");
        getResult.Stdout.Should().Contain("BatchCell");
    }

    [Fact]
    public void Batch_InputFile()
    {
        var path = CreateWorkbook();
        var batchFile = NewTempPath(".json");
        File.WriteAllText(batchFile, """
            [
              {"command":"add","parent":"/Sheet1/A1","type":"cell","props":{"value":"FileBatch"}},
              {"command":"add","parent":"/Sheet1/B1","type":"cell","props":{"value":"Second"}}
            ]
            """);

        var result = RunCliOk("batch", path, "--input", batchFile, "--json");
        result.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", path);
    }

    [Fact]
    public void Batch_CreateSheetAddCells()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("batch", path, "--commands", """
            [
              {"command":"add","parent":"/","type":"sheet","props":{"name":"BatchSheet"}},
              {"command":"add","parent":"/BatchSheet/A1","type":"cell","props":{"value":"Hello"}},
              {"command":"add","parent":"/BatchSheet/B1","type":"cell","props":{"value":"World"}}
            ]
            """, "--json");
        result.Stdout.Should().Contain("\"succeeded\"");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/BatchSheet/A1");
        getResult.Stdout.Should().Contain("Hello");
    }

    // ==================== dump ====================

    [Fact]
    public void Dump_WorkbookToFile()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=DumpTest");

        var dumpPath = NewTempPath(".json");
        var result = RunCliOk("dump", path, "/", "--out", dumpPath);
        File.Exists(dumpPath).Should().BeTrue();
        var dumpContent = File.ReadAllText(dumpPath);
        dumpContent.Should().Contain("DumpTest");
    }

    [Fact]
    public void Dump_ReplayRoundTrip()
    {
        var source = CreateWorkbook();
        RunCliOk("add", source, "/Sheet1/A1", "--type", "cell", "--prop", "value=RoundTrip");

        var dumpPath = NewTempPath(".json");
        RunCliOk("dump", source, "/", "--out", dumpPath);

        var target = CreateWorkbook();
        RunCliOk("batch", target, "--input", dumpPath, "--json");
        RunCliOk("validate", target);

        var getResult = RunCliOk("get", target, "/Sheet1/A1");
        getResult.Stdout.Should().Contain("RoundTrip");
    }

    // ==================== raw ====================

    [Fact]
    public void Raw_WorkbookPart()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("raw", path, "/workbook");
        result.Stdout.Should().Contain("<x:workbook");
    }

    [Fact]
    public void Raw_SheetPart()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=RawSheetTest");
        var result = RunCliOk("raw", path, "/Sheet1");
        result.Stdout.Should().Contain("RawSheetTest");
    }

    // ==================== raw-set ====================

    [Fact]
    public void RawSet_AppendRow()
    {
        var path = CreateWorkbook();
        // Append a new row with a cell via raw XML
        var result = RunCliOk("raw-set", path, "/Sheet1",
            "--xpath", "/x:worksheet/x:sheetData",
            "--action", "append",
            "--xml", "<x:row r=\"99\"><x:c r=\"A99\" t=\"inlineStr\"><x:is><x:t>RawSetValue</x:t></x:is></x:c></x:row>");
        result.Stdout.Should().Contain("raw-set");
    }

    // ==================== validate ====================

    [Fact]
    public void Validate_NewWorkbook_Passes()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("validate", path);
        result.Stdout.Should().Contain("no errors");
    }

    [Fact]
    public void Validate_AfterMutations_Passes()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=ValidSheet");
        RunCliOk("add", path, "/ValidSheet/A1", "--type", "cell", "--prop", "value=Data");

        var result = RunCliOk("validate", path);
        result.Stdout.Should().Contain("no errors");
    }

    // ==================== import ====================

    [Fact]
    public void Import_CsvFile()
    {
        var path = CreateWorkbook();
        var csvPath = NewTempPath(".csv");
        File.WriteAllText(csvPath, "Name,Score\nAlice,95\nBob,87\n");

        var result = RunCliOk("import", path, "/Sheet1", "--file", csvPath, "--header");
        result.Stdout.Should().Contain("Imported");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/Sheet1/A1");
        getResult.Stdout.Should().Contain("Name");
    }

    [Fact]
    public void Import_CsvWithStartCell()
    {
        var path = CreateWorkbook();
        var csvPath = NewTempPath(".csv");
        File.WriteAllText(csvPath, "X,Y\n1,2\n");

        var result = RunCliOk("import", path, "/Sheet1", "--file", csvPath, "--start-cell", "C5");
        result.Stdout.Should().Contain("Imported");
        RunCliOk("validate", path);

        var getResult = RunCliOk("get", path, "/Sheet1/C5");
        getResult.Stdout.Should().Contain("X");
    }

    // ==================== Error cases ====================

    [Fact]
    public void Error_OpenMissingFile()
    {
        var result = RunCli("open", "nonexistent_file_xyz123.xlsx");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_GetInvalidSheet()
    {
        var path = CreateWorkbook();
        var result = RunCli("get", path, "/NoSuchSheet");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_GetInvalidCell()
    {
        var path = CreateWorkbook();
        // Accessing a cell on a non-existent sheet
        var result = RunCli("get", path, "/NoSheet/A1");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_AddInvalidType()
    {
        var path = CreateWorkbook();
        var result = RunCli("add", path, "/Sheet1", "--type", "slide");
        result.ExitCode.Should().NotBe(0);
        (result.Stdout + result.Stderr).Should().Contain("Invalid element type");
    }

    [Fact]
    public void Error_AddInvalidType_Paragraph()
    {
        var path = CreateWorkbook();
        var result = RunCli("add", path, "/Sheet1", "--type", "paragraph");
        result.ExitCode.Should().NotBe(0);
        (result.Stdout + result.Stderr).Should().Contain("Invalid element type");
    }

    [Fact]
    public void Error_RemoveNonexistentSheet()
    {
        var path = CreateWorkbook();
        var result = RunCli("remove", path, "/NoSuchSheet");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_SetInvalidPath()
    {
        var path = CreateWorkbook();
        var result = RunCli("set", path, "/NoSheet/A1", "--prop", "value=Test");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_ValidateMissingFile()
    {
        var result = RunCli("validate", "nonexistent_file_xyz789.xlsx");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_ImportMissingSourceFile()
    {
        var path = CreateWorkbook();
        var result = RunCli("import", path, "/Sheet1", "--file", "nonexistent_csv_xyz.csv");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_AddWithoutType()
    {
        var path = CreateWorkbook();
        var result = RunCli("add", path, "/Sheet1/A1");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_DuplicateSheetName()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=DupSheet");
        var result = RunCli("add", path, "/", "--type", "sheet", "--prop", "name=DupSheet");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_DumpMissingFile()
    {
        var result = RunCli("dump", "nonexistent_file_xyz_dump.xlsx", "/");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_RawSetInvalidPart()
    {
        var path = CreateWorkbook();
        // Try to raw-set on a non-existent part
        var result = RunCli("raw-set", path, "/NoSuchPart",
            "--xpath", "//x:row",
            "--action", "append",
            "--xml", "<x:row r=\"1\"/>");
        result.ExitCode.Should().NotBe(0);
    }

    // ==================== Comprehensive mutation validation ====================

    [Fact]
    public void AllMutations_LeaveWorkbookValid()
    {
        var path = CreateWorkbook();

        // add sheet
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=Data");
        // add cells
        RunCliOk("add", path, "/Data/A1", "--type", "cell", "--prop", "value=Item", "--prop", "bold=true");
        RunCliOk("add", path, "/Data/B1", "--type", "cell", "--prop", "value=Qty");
        RunCliOk("add", path, "/Data/A2", "--type", "cell", "--prop", "value=Pen", "--prop", "font.color=blue");
        RunCliOk("add", path, "/Data/B2", "--type", "cell", "--prop", "value=42", "--prop", "type=number");
        // add formula
        RunCliOk("add", path, "/Data/B3", "--type", "cell", "--prop", "formula=SUM(B2:B2)");
        // set a cell
        RunCliOk("set", path, "/Data/A1", "--prop", "value=Product");
        // add a row
        RunCliOk("add", path, "/Data", "--type", "row", "--prop", "cols=2", "--prop", "c1=New", "--prop", "c2=100");
        // add a named range
        RunCliOk("add", path, "/", "--type", "namedrange", "--prop", "name=MyNR", "--prop", "refersTo=Data!$A$1:$B$2");
        // add a table
        RunCliOk("add", path, "/Data", "--type", "table",
            "--prop", "ref=A1:B2", "--prop", "name=DataTable", "--prop", "columns=Product,Qty");

        // Validate final state
        var validate = RunCliOk("validate", path);
        validate.Stdout.Should().Contain("no errors");
    }

    // ==================== get with --json ====================

    [Fact]
    public void Get_Sheet_JsonOutput()
    {
        var path = CreateWorkbook();
        var result = RunCliOk("get", path, "/Sheet1", "--json");
        result.Stdout.Should().Contain("\"type\"").And.Contain("\"sheet\"");
    }

    // ==================== query edge cases ====================

    [Fact]
    public void Query_Chart()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/Sheet1/A1", "--type", "cell", "--prop", "value=10", "--prop", "type=number");
        RunCliOk("add", path, "/Sheet1/B1", "--type", "cell", "--prop", "value=20", "--prop", "type=number");
        RunCliOk("add", path, "/Sheet1", "--type", "chart",
            "--prop", "chartType=column",
            "--prop", "dataRange=Sheet1!A1:B1",
            "--prop", "anchor=F2:L16",
            "--prop", "title=QChart");

        var result = RunCliOk("query", path, "chart");
        result.Stdout.Should().Contain("QChart");
    }

    [Fact]
    public void Query_NamedRange()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/", "--type", "namedrange",
            "--prop", "name=QNR", "--prop", "refersTo=Sheet1!$A$1:$C$5");

        var result = RunCliOk("query", path, "namedrange");
        result.Stdout.Should().Contain("QNR");
    }

    // ==================== move edge cases ====================

    [Fact]
    public void Move_SheetAfterSheet()
    {
        var path = CreateWorkbook();
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=S1");
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=S2");
        RunCliOk("add", path, "/", "--type", "sheet", "--prop", "name=S3");

        var result = RunCliOk("move", path, "/S1", "--after", "/S3");
        result.Stdout.Should().Contain("Moved");
        RunCliOk("validate", path);
    }

    // ==================== more error cases ====================

    [Fact]
    public void Error_Move_NoTarget()
    {
        var path = CreateWorkbook();
        // Move a sheet without specifying index/after/before
        var result = RunCli("move", path, "/Sheet1");
        result.ExitCode.Should().NotBe(0);
    }

    [Fact]
    public void Error_InvalidType()
    {
        // The CLI should produce a help message or error for unknown subcommands
        var result = RunCli("nonexistent_command_xyz", "somefile.xlsx");
        result.ExitCode.Should().NotBe(0);
    }
}
