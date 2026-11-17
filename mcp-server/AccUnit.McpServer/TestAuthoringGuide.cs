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

To write tests for a module, start with accunit_generate_test_class. Do not read the source of the module under test
first, neither the exported source files nor through other tools.

Use the skill accunit-tests for writing and running tests. If it is not available in the project, offer to install it
with accunit_install_skill (target ""claude"" for Claude Code, ""agents"" for Codex, Copilot and Gemini).

Essential rules if the skill is not available:
- For methods with parameters prefer ONE row test with several 'AccUnit:Row(inputs..., expected).Name(""case"") lines
  above the test Sub over one test per case.
- Row values are VB.NET expressions, not VBA: no constants or enums of the own VBA project (use the numeric value),
  no VBA functions (dates: System.Convert.ToDateTime(""2014-05-16"")); arrays next to other values as {""a"", ""b""},
  Array(...) only as the only value of a row.
- Iz.EqualTo distinguishes vbNullString and """".
- The specification is the method name plus the declaration and comment returned by accunit_generate_test_class;
  do not read the body of the method under test on your own. If the expected behavior cannot be derived, ask the
  developer once how to proceed (assume values, look at the code, developer specifies, leave out).
- Expected error: Assert.Throws with the VBA error number as seen in Err.Number (e.g. Assert.Throws 457).
- Save the test class as <TestClassName>.cls in the testFolder returned by accunit_generate_test_class and import it
  with accunit_import_test_classes.";
    }
}
