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
        private readonly TestClassImportService _testClassImportService;
        private readonly TestModuleExportService _testModuleExportService;

        public TestTools(TestRunService testRunService, TestClassDraftService testClassDraftService,
                         TestEnvironmentService testEnvironmentService, TestClassImportService testClassImportService,
                         TestModuleExportService testModuleExportService)
        {
            _testRunService = testRunService;
            _testClassDraftService = testClassDraftService;
            _testEnvironmentService = testEnvironmentService;
            _testClassImportService = testClassImportService;
            _testModuleExportService = testModuleExportService;
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
                     "Save the finished class as a source file (e.g. <ClassName>.cls in the AccUnit test folder) and import it with " +
                     "accunit_import_test_classes. If a test class for the module exists already, the proposed name gets a number " +
                     "(e.g. modCalcTests1): create, import and run it, then ask the developer whether to merge it into the existing class, " +
                     "replace the existing class or keep both.")]
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

        [McpServerTool(Name = "accunit_import_test_classes", Destructive = true)]
        [Description("Imports test classes and test related modules from source files (.cls/.bas, VBA export format, UTF-8 or ANSI) " +
                     "into the open database with the import of the AccUnit core, as the VBE add-in does. Only files with the header comment " +
                     "'AccUnit:TestClass or 'AccUnit:TestRelated are imported; the file name must match the module name. An existing test module " +
                     "with the same name is replaced; a module that is not a test module is never replaced. " +
                     "Default folder: the AccUnit test folder of the database (<database folder>\\Tests\\<database name>, returned as importDirectory). " +
                     "To keep the import, Access is closed with saving and the database is opened again.")]
        public TestClassImportResult ImportTestClasses(
            [Description("File name without extension, wildcards * and ? allowed (e.g. modStringsTests).")] string fileNameFilter,
            [Description("Absolute path of the folder with the source files (optional, default: AccUnit test folder of the database).")] string importPath = null)
        {
            try
            {
                return _testClassImportService.ImportTestClasses(fileNameFilter, importPath);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }

        [McpServerTool(Name = "accunit_export_test_classes")]
        [Description("Exports test classes and test related modules of the open database to the AccUnit test folder (VBA export " +
                     "format, ANSI), e.g. to edit an existing test class and import it again with accunit_import_test_classes. " +
                     "Only modules with 'AccUnit:TestClass or 'AccUnit:TestRelated are exported. An existing file is written over only " +
                     "with overwrite = true and is kept as .bak. Nothing is written to the database.")]
        public ExportTestClassesResult ExportTestClasses(
            [Description("Names of the test classes or test related modules.")] string[] classNames,
            [Description("Write over an existing file in the test folder (kept as .bak).")] bool overwrite = false)
        {
            try
            {
                return _testModuleExportService.ExportTestClasses(classNames, overwrite);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }

        [McpServerTool(Name = "accunit_remove_test_classes", Destructive = true)]
        [Description("Removes single test classes or test related modules from the open database, e.g. a second test class after it " +
                     "was merged into the existing one. Only modules with 'AccUnit:TestClass or 'AccUnit:TestRelated are removed; " +
                     "they are exported to the AccUnit test folder first by default (an existing file is kept as .bak). " +
                     "The module AccUnit_TestClassFactory is removed as well; accunit_run_tests adds it again. " +
                     "To keep the removal, Access is closed with saving and the database is opened again.")]
        public RemoveTestClassesResult RemoveTestClasses(
            [Description("Names of the test classes or test related modules.")] string[] classNames,
            [Description("Export the modules to the AccUnit test folder before removing them.")] bool exportBeforeRemoving = true)
        {
            try
            {
                return _testModuleExportService.RemoveTestClasses(classNames, exportBeforeRemoving);
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
