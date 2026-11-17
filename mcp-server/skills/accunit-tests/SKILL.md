---
name: accunit-tests
description: Write and run AccUnit unit tests for VBA code in Microsoft Access - test classes, row tests ('AccUnit:Row), Assert/Iz - with the AccUnit MCP server (accunit_* tools). Use when the user asks to create unit tests for VBA code, when creating or changing AccUnit test classes or when running AccUnit tests.
---

# AccUnit tests

AccUnit runs unit tests for VBA code in Microsoft Access. The AccUnit MCP server (tools `accunit_*`) lists and runs
the tests, creates drafts of test classes and imports test classes into the database.

## Workflow

1. `accunit_open` the database (absolute path; it must be in a trusted location).
2. `accunit_generate_test_class` for the module under test. Start here: do not read the source of the module under
   test first, neither the exported source files nor through other tools. You get a draft plus the declaration and the comment of
   each method under test, read from the current code in the database. Methods with parameters already get a row test skeleton.
3. Fill the tests (see below) and save the class as `<TestClassName>.cls` in the `testFolder` returned by the draft
   (the AccUnit test folder of the database). The draft code can be saved as it is; the export header is added on import.
4. `accunit_import_test_classes` with the class name as `fileNameFilter` (optionally `importPath` for another folder,
   e.g. the source folder of msaccess-vcs). Only modules with `'AccUnit:TestClass` or `'AccUnit:TestRelated` are imported;
   an existing test module with the same name is replaced. Import again after each change of the file.
5. `accunit_run_tests` (optionally `classNames`, `testNameFilter`, `tags`); it returns a summary and the failed,
   error and ignored tests. Fix the file, import and run again.
6. `accunit_close` when done.

Notes on running:
- The tests run inside Access with the installed AccUnitLoader add-in; private test classes work unchanged.
- The server inserts the module `AccUnit_Factory` and factory methods in `AccUnit_TestClassFactory` if needed and
  lists that in `environmentChanges`; these modules belong to the test environment, do not edit them.
- "The VBA project does not compile": ask the developer to compile in the VBA editor (Debug > Compile) and to report the line.
- Never run tests in parallel against the same database.

## Writing tests

- Test class: class module named `<ModuleUnderTest>Tests` with the comment `'AccUnit:TestClass`,
  `Option Compare Text`, `Option Explicit`. Every public `Sub` is a test; `Setup` and `Teardown` run before/after each test.
- The specification is the method name plus the declaration and comment returned by
  `accunit_generate_test_class`. Do not read the body of the method under test on your own, neither in the
  exported source files nor through other tools; this includes the procedures it calls. Public declarations
  used by the method (enums, constants) may be looked up.
  If the expected behavior cannot be derived from that for some methods or cases (e.g. Null, empty
  input, boundaries, error behavior), tell the developer once, listing the affected methods/cases
  briefly, and let them choose how to proceed:
  - **Assume values:** choose expected values from the method's purpose and common conventions; mark
    them with `' assumption: ...` and list them in the summary.
  - **Look at the code:** derive the expected values from the implementation; the tests then pin the
    current behavior (characterization tests). Say so in the summary.
  - **Developer specifies:** the developer states the expected behavior.
  - **Leave out:** skip these cases.
  A failing test is a result: report it as "assumption or code wrong?" and do not adjust the
  expectation to the current behavior unless the developer agrees.
- One behavior per test; Arrange - Act - Assert.
- Assertions: `Assert.That Actual, Iz.EqualTo(Expected)` (also `Iz.Empty`, `Iz.Null`, ...).
  Expected error: `Assert.Throws <ErrNumber>` before the call that raises it, with the VBA error number
  as seen in `Err.Number` (e.g. `Assert.Throws 457`).
- Names: `MethodUnderTest_StateUnderTest_ExpectedBehaviour` is recommended; leave out parts that add nothing
  (e.g. `NetPrice_ReturnsZeroForZero`).
- Annotations above a test: `'AccUnit:Tags(Fast, Database)`, `'AccUnit:Ignore reason text`
  (only for tests whose expected behavior is still open; for expected errors use `Assert.Throws`),
  `'AccUnit:Rollback` (the test changes data).

## Row tests

For methods with parameters prefer ONE test method with several `'AccUnit:Row` lines over one test method per case
(normal case, boundaries, invalid input, Null). Example:

```vba
'AccUnit:Row(2, "a", "F = 'a'").Name("Equal")
'AccUnit:Row(2+8, "a", "F >= 'a'").Name("Equal+GreaterThan")
'AccUnit:Row(2, Null, "").Name("Null value")
'AccUnit:Row(2, System.Convert.ToDateTime("2014-05-16"), "F = #2014-05-16#").Name("Date")
Public Sub BuildCriteria_ReturnsCriteria(ByVal RelationalOperator As Long, ByVal FilterValue As Variant, ByVal Expected As String)
    Dim Actual As String
    Actual = modCriteria.BuildCriteria("F", RelationalOperator, FilterValue)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
```

Rules for row values - AccUnit compiles them as VB.NET expressions, not VBA:
- One value per parameter of the test method, in parameter order: inputs first, expected value last.
- Constants and enums of the own VBA project are not available: use the numeric value (`2+8` above), add a comment if helpful.
- VBA functions (`DateSerial`, `CDate`, `Chr`, ...) are not available: use .NET expressions,
  e.g. `System.Convert.ToDateTime("2014-05-16")`.
- Numbers with decimal point; strings in double quotes (double a quote inside a string); `Null` passes VBA Null;
  VBA constants such as `vbTextCompare` work.
- Arrays next to other values in VB.NET syntax: `'AccUnit:Row(1, {"a", "b"}, ...)`.
  `Array("a", "b")` only when the array is the only value of the row; next to other values it arrives nested.
- Whole numbers arrive as Long (VB.NET Integer), so use `Long` or `Variant` parameters.
- `.Name("case")` for each row makes the results readable.

Strings: `Iz.EqualTo` distinguishes `vbNullString` and `""`. A String function that never assigns its result returns
`vbNullString`; compare `"" & Actual` or use `Iz.StringCompare(StringCompareMode_vbNullStringEqualEmptyString)`.
