// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

Console.OutputEncoding = System.Text.Encoding.UTF8;

OfficeCli.Core.LocaleFontRegistry.OsLocaleSnapshot =
    System.Globalization.CultureInfo.CurrentCulture.Name;
System.Globalization.CultureInfo.DefaultThreadCurrentCulture =
    System.Globalization.CultureInfo.InvariantCulture;
var uiCulture = System.Globalization.CultureInfo.GetCultureInfo("zh-Hans");
System.Globalization.CultureInfo.CurrentUICulture = uiCulture;
System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = uiCulture;

// Keep --help and help on one surface without rewriting --help used as a value.
if (args.Length > 0)
{
    if (args[0] is "--help" or "-h" or "-?")
    {
        var tail = args.Skip(1).ToArray();
        args = tail.Length == 0 ? ["help"] : ["help", .. tail];
    }
    else if (args.Length >= 2 && args[1] is "--help" or "-h" or "-?")
    {
        var tail = args.Skip(2).ToArray();
        args = tail.Length == 0 ? ["help", args[0]] : ["help", args[0], .. tail];
    }
}

// MCP remains available both as a stdio server and through its existing
// explicit registration commands.
if (args.Length >= 1 && args[0] == "mcp")
{
    if (args.Length == 1)
    {
        await OfficeCli.McpServer.RunAsync();
        return 0;
    }
    if (args.Length == 2 && args[1] == "list")
    {
        OfficeCli.McpInstaller.Install("list");
        return 0;
    }
    if (args.Length == 3 && args[1] == "uninstall")
        return OfficeCli.McpInstaller.Uninstall(args[2]) ? 0 : 1;
    if (args.Length == 2)
        return OfficeCli.McpInstaller.Install(args[1]) ? 0 : 1;

    OfficeCli.CommandBuilder.WriteEarlyDispatchUsage("mcp", Console.Error);
    return 1;
}

if (args.Length == 1 && args[0] == "mcp-serve")
{
    await OfficeCli.McpServer.RunAsync();
    return 0;
}

// Read-only access to embedded guidance remains available to both CLI and MCP.
if (args.Length >= 1 && args[0] == "load_skill")
{
    string? skillRelPath = null;
    var positional = new List<string>();
    for (var i = 1; i < args.Length; i++)
    {
        if (args[i] == "--path" && i + 1 < args.Length)
        {
            skillRelPath = args[++i];
            continue;
        }
        positional.Add(args[i]);
    }

    if (positional.Count == 0 && string.IsNullOrEmpty(skillRelPath))
    {
        Console.Out.Write(OfficeCli.Core.SkillCatalog.BuildSkillCatalog());
        return 0;
    }
    if (positional.Count == 1)
    {
        try
        {
            Console.Out.Write(string.IsNullOrEmpty(skillRelPath)
                ? OfficeCli.Core.SkillCatalog.LoadSkillContent(positional[0])
                : OfficeCli.Core.SkillCatalog.LoadSkillFile(positional[0], skillRelPath));
            return 0;
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }

    OfficeCli.CommandBuilder.WriteEarlyDispatchUsage("load_skill", Console.Error);
    return 1;
}

var rootCommand = OfficeCli.CommandBuilder.BuildRootCommand();
if (args.Length == 0)
{
    rootCommand.Parse("help").Invoke();
    return 0;
}

var parseResult = rootCommand.Parse(args,
    new System.CommandLine.ParserConfiguration { ResponseFileTokenReplacer = null });
return parseResult.Invoke();
