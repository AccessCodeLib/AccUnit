---
name: accunit-tests
description: Write and run AccUnit unit tests for VBA code in Microsoft Access - test classes, row tests ('AccUnit:Row), Assert/Iz - with the AccUnit MCP server (accunit_* tools). Use when creating or changing AccUnit test classes or when running AccUnit tests.
---

# AccUnit tests

AccUnit runs unit tests for VBA code in Microsoft Access. The AccUnit MCP server (tools `accunit_*`) lists and runs
the tests and creates drafts of test classes.

## Workflow

1. `accunit_open` the database (absolute path; it must be in a trusted location).
2. `accunit_generate_test_class` for the module under test: you get a draft plus the declaration and the comment of
   each method under test, read from the current code in the database. Methods with parameters already get a row test skeleton.
3. Fill the tests (see below) and add the test class through the source files (e.g. msaccess-vcs export/import).
   The server does not write test code into the database.
4. `accunit_run_tests` (optionally `classNames`, `testNameFilter`, `tags`); it returns a summary and the failed,
   error and ignored tests. Fix and run again.
5. `accunit_close` when done.

Notes on running:
- The tests run inside Access with the installed AccUnitLoader add-in; private test classes work unchanged.
- The server inserts the module `AccUnit_Factory` and factory methods in `AccUnit_TestClassFactory` if needed and
  lists that in `environmentChanges`; these modules belong to the test environment, do not edit them.
- "The VBA project does not compile": ask the developer to compile in the VBA editor (Debug > Compile) and to report the line.
- Never run tests in parallel against the same database.

## Writing tests

- Test class: class module named `<ModuleUnderTest>Tests` with the comment `'AccUnit:TestClass`,
  `Option Compare Text`, `Option Explicit`. Every public `Sub` is a test; `Setup` and `Teardown` run before/after each test.
- Write the tests from the specification or the task, not from the implementation, so they do not just confirm the
  current behavior. If the specification is unclear, ask the developer.
- One behavior per test; Arrange - Act - Assert.
- Assertions: `Assert.That Actual, Iz.EqualTo(Expected)` (also `Iz.Empty`, `Iz.Null`, ...).
  Expected error: `Assert.Throws <ErrNumber>` before the call that raises it.
- Names: `MethodUnderTest_StateUnderTest_ExpectedBehaviour` is recommended; leave out parts that add nothing
  (e.g. `NetPrice_ReturnsZeroForZero`).
- Annotations above a test: `'AccUnit:Tags(Fast, Database)`, `'AccUnit:Ignore reason text`,
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
- `Array(...)` works only as the only value of a row; together with other values the method receives a nested array.
- `.Name("case")` for each row makes the results readable.

Strings: `Iz.EqualTo` distinguishes `vbNullString` and `""`. A String function that never assigns its result returns
`vbNullString`; compare `"" & Actual` or use `Iz.StringCompare(StringCompareMode_vbNullStringEqualEmptyString)`.
