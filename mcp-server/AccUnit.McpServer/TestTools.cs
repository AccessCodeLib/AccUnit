using System.Collections.Generic;
using System.ComponentModel;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace AccessCodeLib.AccUnit.McpServer
{
    [McpServerToolType]
    public class TestTools
    {
        private readonly TestRunService _testRunService;
        private readonly TestClassDraftService _testClassDraftService;
        private readonly TestEnvironmentService _testEnvironmentService;

        public TestTools(TestRunService testRunService, TestClassDraftService testClassDraftService, TestEnvironmentService testEnvironmentService)
        {
            _testRunService = testRunService;
            _testClassDraftService = testClassDraftService;
            _testEnvironmentService = testEnvironmentService;
        }

        [McpServerTool(Name = "accunit_generate_test_class", ReadOnly = true)]
        [Description("Creates the DRAFT of an AccUnit test class for a module of the open database (same generator as the VBE add-in): " +
                     "class header with 'AccUnit:TestClass and one empty test method per public method, plus the declaration and the comment " +
                     "(e.g. NaturalDoc) of each method under test, read from the current code in the database. Nothing is written to the database. " +
                     "Fill the test methods (Arrange - Act - Assert) from the specification: the method name plus the returned declaration " +
                     "and comment. Do not read the body of the method under test on your own, so the tests do not just confirm the current behavior. " +
                     "If the expected behavior cannot be derived for some methods or cases, ask the developer once how to proceed " +
                     "(assume values, look at the code, developer specifies, leave out). " +
                     "Test names: the recommended pattern is MethodUnderTest_StateUnderTest_ExpectedBehaviour (replace the placeholders); " +
                     "it is a recommendation, leave out parts that add nothing (e.g. NetPrice_ReturnsZeroForZero). " +
                     "Methods with parameters get a row test skeleton: write one 'AccUnit:Row(inputs..., expected).Name(\"case\") line per case " +
                     "(normal case, boundaries, invalid input, Null) instead of separate test methods. Row values are VB.NET expressions, not VBA: " +
                     "no constants or enums of the own VBA project (use the numeric value), no VBA functions (dates: " +
                     "System.Convert.ToDateTime(\"2014-05-16\")); see the server instructions for an example. " +
                     "Add the finished class through the source files (e.g. msaccess-vcs import), not by writing into the database.")]
        public TestClassDraft GenerateTestClass(
            [Description("Name of the standard module or class module under test.")] string moduleName,
            [Description("Only these public methods (optional, default: all public methods).")] string[] methodNames = null)
        {
            try
            {
                return _testClassDraftService.CreateDraft(moduleName, methodNames);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }

        [McpServerTool(Name = "accunit_remove_test_environment", Destructive = true)]
        [Description("Removes the AccUnit test environment from the open database like the menu item of the VBE add-in: " +
                     "the modules AccUnit_Factory and AccUnit_TestClassFactory and the AccUnit type library reference. " +
                     "With removeTestModules = true the test classes are removed as well (exported to a folder first by default). " +
                     "To keep the removal, Access is closed with saving and the database is opened again.")]
        public RemoveTestEnvironmentResult RemoveTestEnvironment(
            [Description("Also remove the test classes and test related modules.")] bool removeTestModules = false,
            [Description("Export the test classes before removing them.")] bool exportBeforeRemoving = true)
        {
            try
            {
                return _testEnvironmentService.RemoveTestEnvironment(removeTestModules, exportBeforeRemoving);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }

        [McpServerTool(Name = "accunit_list_tests", ReadOnly = true)]
        [Description("Lists the AccUnit test classes of the open database with their test methods, tags, ignore info and number of data rows (row tests).")]
        public List<TestClassListing> ListTests(
            [Description("Only these test classes (optional).")] string[] classNames = null)
        {
            try
            {
                return _testRunService.ListTests(classNames);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }

        [McpServerTool(Name = "accunit_run_tests")]
        [Description("Runs AccUnit tests in the open database and returns a summary plus all failed, error and ignored tests. " +
                     "The tests run inside Access with the test suite of the installed AccUnitLoader add-in, as TestSuite in the Immediate window; " +
                     "the results are read from the JUnit XML file AccUnit writes. If needed the server inserts the module AccUnit_Factory and " +
                     "adds factory methods to AccUnit_TestClassFactory; these changes are kept and listed in environmentChanges.")]
        public TestRunResult RunTests(
            [Description("Only these test classes (optional, default: all test classes).")] string[] classNames = null,
            [Description("Only tests whose name matches one of these filters, wildcards * and ? allowed (optional).")] string[] testNameFilter = null,
            [Description("Only tests with all of these tags (optional).")] string[] tags = null,
            [Description("Also list passed tests in results.")] bool includePassed = false)
        {
            try
            {
                return _testRunService.RunTests(classNames, testNameFilter, tags, includePassed);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }
    }
}
