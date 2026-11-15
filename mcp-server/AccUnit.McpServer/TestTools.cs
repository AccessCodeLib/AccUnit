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

        public TestTools(TestRunService testRunService)
        {
            _testRunService = testRunService;
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
