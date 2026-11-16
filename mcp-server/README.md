# AccUnit MCP server

MCP server (Model Context Protocol, stdio) that lets an AI agent such as Claude Code work with AccUnit tests
in a Microsoft Access database: list and run tests, create drafts of test classes and remove the test environment.

The tests run inside Access with the test suite of the installed AccUnitLoader add-in,
the same way as `TestSuite.AddByClassName("...").Run` in the VBA Immediate window.

## Requirements

- Windows with Microsoft Access (32 or 64 bit)
- .NET Framework 4.8
- AccUnit installed: the AccUnitLoader add-in `AccUnitLoader.accda` in `%APPDATA%\Microsoft\AddIns`
  (see [How-to: Install AccUnit](https://github.com/AccessCodeLib/AccUnit/wiki/How%E2%80%90to:-Install-AccUnit))
- The database is in a trusted location, otherwise Access does not run its VBA code.

## Installation

1. Download `AccUnit.McpServer.<version>.zip` from the [AccUnit release](https://github.com/AccessCodeLib/AccUnit/releases)
   (same version as AccUnit).
2. Unblock the zip before extracting it (file properties, or `Unblock-File` in PowerShell);
   otherwise Windows may block the DLLs that were downloaded from the internet.
3. Extract it to a local folder, e.g. `C:\Tools\AccUnit.McpServer`. The DLLs must not be on a network drive.

## Configuration

### Claude Code

```
claude mcp add accunit -- "C:\Tools\AccUnit.McpServer\AccUnit.McpServer.exe"
```

or in the `.mcp.json` of a project:

```json
{
  "mcpServers": {
    "accunit": {
      "type": "stdio",
      "command": "C:\\Tools\\AccUnit.McpServer\\AccUnit.McpServer.exe",
      "args": []
    }
  }
}
```

### Other MCP clients

Start `AccUnit.McpServer.exe` as a stdio server without arguments. Log output goes to stderr.

## Tools

| Tool | Parameters | Purpose |
|---|---|---|
| `accunit_open` | `filePath`, `bypassStartup` (default `true`) | Opens the database (.accdb, .accda, .mdb) exclusively in a visible Access instance owned by the server. |
| `accunit_close` | | Closes the database without saving and quits Access. |
| `accunit_status` | | Path, Access version, process id, whether the database is trusted. |
| `accunit_list_tests` | `classNames` | Test classes with test methods, tags, ignore info and number of data rows (row tests). |
| `accunit_run_tests` | `classNames`, `testNameFilter`, `tags`, `includePassed` (default `false`) | Runs the tests; returns a summary and the failed, error and ignored tests. |
| `accunit_generate_test_class` | `moduleName`, `methodNames` | Draft of a test class for a module: code skeleton plus declaration and comment of each method under test. Writes nothing to the database. |
| `accunit_remove_test_environment` | `removeTestModules` (default `false`), `exportBeforeRemoving` (default `true`) | Removes the test environment like the VBE add-in menu item; optionally also the test classes (exported first). |
| `accunit_install_skill` | `target` (`claude`/`agents`), `scope` (`project`/`user`), `projectPath`, `overwrite` (default `false`) | Installs the skill `accunit-tests` shipped with the server (see below). |

### Writing tests with `accunit_generate_test_class`

The tool returns a draft only. The agent fills the test methods (Arrange - Act - Assert) from the specification
or the task, not from the implementation, so the tests do not just confirm the current behavior.
If the specification is unclear, the agent should ask the developer.
The recommended test name pattern is `MethodUnderTest_StateUnderTest_ExpectedBehaviour`; parts that add nothing may be left out.
The finished test class is added through the source files (e.g. msaccess-vcs import), not written into the database by the server.

## Skill `accunit-tests`

The guide for writing and running AccUnit tests is the skill `accunit-tests` (`skills/accunit-tests/SKILL.md`,
portable Agent Skills format). Agents load it only when they work on AccUnit tests. It asks for row tests for methods
with parameters and explains the rules for row values: AccUnit compiles them as VB.NET expressions, so constants and
enums of the own VBA project and VBA functions are not available (use numeric values and .NET expressions such as
`System.Convert.ToDateTime("2014-05-16")`), and `Array(...)` works only as the only value of a row.

Install it with `accunit_install_skill` or copy the folder `skills/accunit-tests` from the zip:

| `target` | `scope` = `project` (default) | `scope` = `user` | Used by |
|---|---|---|---|
| `claude` | `<project>\.claude\skills\accunit-tests` | `%USERPROFILE%\.claude\skills\accunit-tests` | Claude Code (also GitHub Copilot) |
| `agents` | `<project>\.agents\skills\accunit-tests` | `%USERPROFILE%\.agents\skills\accunit-tests` | Codex, GitHub Copilot, Gemini CLI |

With `scope` = `project` the skill is only active in the Access projects that use AccUnit and can be committed with the
project. An installed skill is not overwritten unless `overwrite` = `true` (it may contain local improvements);
changed files are backed up as `.bak` first. Improvements made locally can be brought back into this repository.

The server also sends short MCP server instructions with the essential rules, for clients without skill support
(source: `AccUnit.McpServer/TestAuthoringGuide.cs`).

A larger example of row tests: [SqlToolsBuildCriteriaTests.cls](https://github.com/AccessCodeLib/AccessCodeLib/blob/master/_test/data/SqlToolsBuildCriteriaTests.cls) in the AccessCodeLib repository.

## Behavior and changes to the database

- **One database at a time**, opened exclusively in an Access instance started by the server; the server never uses
  an Access instance it did not start. The instance is visible.
- **Startup code** (AutoExec, startup form) is skipped by default like opening with the Shift key held down.
  This does not work if the database sets `AllowBypassKey = False`; the result then contains a warning.
- **`accunit_run_tests`** inserts the module `AccUnit_Factory` if it is missing (test classes use `Assert`/`Iz` from it)
  and adds missing factory methods to `AccUnit_TestClassFactory`. These changes are kept and listed in
  `environmentChanges`. Test classes are not changed. The results are read from the JUnit XML file AccUnit writes
  (`%TEMP%\AccUnitMcpServer`).
- **`accunit_remove_test_environment`** saves the database and opens it again: Access keeps removed modules only
  when the VBA project is saved.
- Tests never run in parallel; all calls to Access are serialized.

## Limitations

- Microsoft Access only. The tool names are host neutral, so other hosts (e.g. Excel) can be added later.
- The AccUnitLoader add-in must be installed; without it `accunit_run_tests` reports an error and does not change the database.

## Build

Solution `AccessCodeLib.AccUnit.McpServer.sln` (.NET Framework 4.8). It references the AccUnit projects in `../source`;
restore their NuGet packages first (`nuget restore ../source/AccessCodeLib.AccUnit.sln`), then build with MSBuild:

```
msbuild AccessCodeLib.AccUnit.McpServer.sln /restore /p:Configuration=Release
```
