namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Server instructions sent to every MCP client. Kept short: the full guide is the skill accunit-tests
    /// (skills/accunit-tests/SKILL.md). Only the rules that depend on the AccUnit core are repeated here,
    /// so clients without skill support get them as well.
    /// </summary>
    public static class TestAuthoringGuide
    {
        public const string SkillName = "accunit-tests";

        public const string Text =
@"AccUnit MCP server: run and write AccUnit tests for the VBA code of a Microsoft Access database.

Use the skill accunit-tests for writing and running tests. If it is not available in the project, offer to install it
with accunit_install_skill (target ""claude"" for Claude Code, ""agents"" for Codex, Copilot and Gemini).

Essential rules if the skill is not available:
- For methods with parameters prefer ONE row test with several 'AccUnit:Row(inputs..., expected).Name(""case"") lines
  above the test Sub over one test per case.
- Row values are VB.NET expressions, not VBA: no constants or enums of the own VBA project (use the numeric value),
  no VBA functions (dates: System.Convert.ToDateTime(""2014-05-16"")); arrays next to other values as {""a"", ""b""},
  Array(...) only as the only value of a row.
- Iz.EqualTo distinguishes vbNullString and """".
- Write the tests from the specification, not from the implementation; add test classes through the source files.";
    }
}
