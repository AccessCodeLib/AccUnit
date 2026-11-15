using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.Interfaces;
using Microsoft.Vbe.Interop;
using Access = Microsoft.Office.Interop.Access;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Lists and runs the AccUnit tests of the database opened in the <see cref="AccessSession"/>.
    /// </summary>
    public sealed class TestRunService
    {
        private const string ProjectDoesNotCompileMessage =
            "The VBA project does not compile. Open the database in Access and compile it (Debug > Compile) to see the error.";
        private const string LoaderNotAvailableMessage =
            "The tests run with the AccUnitLoader add-in, which was not found. Install AccUnit " +
            "(AccUnitLoader.accda in %APPDATA%\\Microsoft\\AddIns) and run the tests again.";
        private const string LoaderFileName = "AccUnitLoader";
        private const string LoaderFileExtension = ".accda";
        private const string LoaderFactoryProcedureName = "GetAccUnitFactory";
        private const int NoTestReportOutput = 0;   // TestReportOutput: no DebugPrint/LogFile reporter

        private readonly AccessSession _session;

        public TestRunService(AccessSession session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        public List<TestClassListing> ListTests(IList<string> classNames)
        {
            return _session.Execute(application => ListTestsOnStaThread(GetVBProject(application), classNames));
        }

        public TestRunResult RunTests(IList<string> classNames, IList<string> testNameFilter, IList<string> tags, bool includePassed)
        {
            return _session.Execute(application =>
            {
                try
                {
                    return RunTestsOnStaThread(application, classNames, testNameFilter, tags, includePassed);
                }
                finally
                {
                    // Release the COM objects of the loader and the test suite now; Access does not quit while they are referenced.
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
        }

        private static List<TestClassListing> ListTestsOnStaThread(VBProject vbProject, IList<string> classNames)
        {
            var reader = new TestClassReader(vbProject);
            return reader.GetTestClasses(true)
                .Where(testClass => IsSelected(testClass.Name, classNames))
                .Select(testClass => new TestClassListing
                {
                    ClassName = testClass.Name,
                    Tags = GetTagNames(testClass.Tags),
                    Tests = testClass.Members.Select(member => new TestMethodListing
                    {
                        Name = member.Name,
                        Tags = GetTagNames(member.Tags),
                        Ignored = member.IgnoreInfo.Ignore,
                        IgnoreComment = string.IsNullOrEmpty(member.IgnoreInfo.Comment) ? null : member.IgnoreInfo.Comment,
                        // GetTestClasses does not read the data rows; only the member info of a single member does.
                        RowCount = reader.GetTestClassMemberInfo(testClass.Name, member.Name).TestRows.Count
                    }).ToList()
                })
                .ToList();
        }

        private static TestRunResult RunTestsOnStaThread(Access.Application application, IList<string> classNames,
                                                         IList<string> testNameFilter, IList<string> tags, bool includePassed)
        {
            // Check first, so the database is not changed when no test can run.
            EnsureLoaderIsInstalled();

            var vbProject = GetVBProject(application);
            var environment = new TestEnvironment(application, vbProject);

            // Test classes use Assert/Iz from AccUnit_Factory, so the project only compiles with that module.
            environment.EnsureFactoryModule();
            if (!environment.Compile())
                throw new DatabaseSessionException(ProjectDoesNotCompileMessage);

            var notRunnable = new List<NotRunnableTestClass>();
            var selectedClasses = SelectTestClasses(vbProject, classNames, notRunnable);

            var testResults = new List<TestResultInfo>();
            if (selectedClasses.Count > 0)
            {
                environment.EnsureTestClassFactoryMethods(selectedClasses);
                if (!environment.Compile())
                    throw new DatabaseSessionException(ProjectDoesNotCompileMessage);
                testResults = RunTestSuite(application, selectedClasses, testNameFilter, tags);
            }

            var result = TestResultMapper.CreateRunResult(testResults, notRunnable, includePassed);
            result.EnvironmentChanges.AddRange(environment.Changes);
            return result;
        }

        private static List<string> SelectTestClasses(VBProject vbProject, IList<string> classNames, List<NotRunnableTestClass> notRunnable)
        {
            var testClassNames = new TestClassReader(vbProject).GetTestClasses().Select(testClass => testClass.Name).ToList();
            if (classNames == null || classNames.Count == 0)
                return testClassNames;

            foreach (var missing in classNames.Where(name => !testClassNames.Contains(name, StringComparer.OrdinalIgnoreCase)))
            {
                notRunnable.Add(new NotRunnableTestClass
                {
                    ClassName = missing,
                    Reason = "No test class with this name. A test class needs the header comment: 'AccUnit:TestClass"
                });
            }
            return testClassNames.Where(name => IsSelected(name, classNames)).ToList();
        }

        /// <summary>
        /// Same as TestSuite.AddByClassName("...").Run in the Immediate window, but with the test suite of the
        /// AccUnitLoader factory that also writes a JUnit XML file: the suite runs inside Access, so private test
        /// classes work, and the results are read from the file (the test objects are not reachable late bound).
        /// </summary>
        private static List<TestResultInfo> RunTestSuite(Access.Application application, IEnumerable<string> classNames,
                                                         IList<string> testNameFilter, IList<string> tags)
        {
            var junitXmlFilePath = PrepareJUnitXmlFile();
            var testSuite = LateBinding.GetProperty(GetLoaderFactory(application), "TestSuite", NoTestReportOutput, junitXmlFilePath);

            foreach (var className in classNames)
                LateBinding.CallMethod(testSuite, "AddByClassName", className);
            if (testNameFilter != null && testNameFilter.Count > 0)
                LateBinding.CallMethod(testSuite, "SelectTests", testNameFilter.ToArray());
            if (tags != null && tags.Count > 0)
                LateBinding.CallMethod(testSuite, "Filter", tags.ToArray());

            LateBinding.CallMethod(testSuite, "Run");

            if (!File.Exists(junitXmlFilePath))
                throw new DatabaseSessionException("AccUnit did not write the test results (" + junitXmlFilePath + ").");
            return JUnitResultReader.Read(File.ReadAllText(junitXmlFilePath));
        }

        private static string LoaderPath
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                    "Microsoft", "AddIns", LoaderFileName);
            }
        }

        private static void EnsureLoaderIsInstalled()
        {
            if (!File.Exists(LoaderPath + LoaderFileExtension))
                throw new DatabaseSessionException(LoaderNotAvailableMessage);
        }

        private static object GetLoaderFactory(Access.Application application)
        {
            try
            {
                // Application.Run has 30 ByRef parameters in the interop assembly; call it late bound.
                return LateBinding.CallMethod(application, "Run", LoaderPath + "." + LoaderFactoryProcedureName);
            }
            catch (COMException ex)
            {
                throw new DatabaseSessionException(LoaderNotAvailableMessage + " (" + ex.Message + ")", ex);
            }
        }

        private static string PrepareJUnitXmlFile()
        {
            var directory = Path.Combine(Path.GetTempPath(), "AccUnitMcpServer");
            Directory.CreateDirectory(directory);
            var filePath = Path.Combine(directory, "TestResults-" + Process.GetCurrentProcess().Id + ".xml");
            // A result file of an earlier run must never be read as the result of this run.
            if (File.Exists(filePath))
                File.Delete(filePath);
            return filePath;
        }

        /// <summary>
        /// VBE.ActiveVBProject can point to a loaded library (e.g. a protected Access wizard) after code ran,
        /// so the project is found by the file name of the open database.
        /// </summary>
        private static VBProject GetVBProject(Access.Application application)
        {
            var databasePath = GetCurrentProjectFullName(application);
            var vbe = (VBE)application.VBE;
            foreach (VBProject vbProject in vbe.VBProjects)
            {
                if (string.Equals(GetFileNameOrNull(vbProject), databasePath, StringComparison.OrdinalIgnoreCase))
                    return vbProject;
            }
            throw new DatabaseSessionException("The VBA project of the open database was not found: " + databasePath);
        }

        private static string GetCurrentProjectFullName(Access.Application application)
        {
            var currentProject = application.CurrentProject;
            try
            {
                return currentProject.FullName;
            }
            finally
            {
                Marshal.ReleaseComObject(currentProject);
            }
        }

        private static string GetFileNameOrNull(VBProject vbProject)
        {
            try
            {
                return vbProject.FileName;
            }
            catch (COMException)
            {
                // Protected or not yet saved projects do not expose their file name.
                return null;
            }
        }

        private static bool IsSelected(string className, IList<string> classNames)
        {
            return classNames == null || classNames.Count == 0 ||
                   classNames.Contains(className, StringComparer.OrdinalIgnoreCase);
        }

        private static List<string> GetTagNames(ITagList tags)
        {
            return tags == null ? new List<string>() : tags.Select(tag => tag.Name).ToList();
        }
    }
}
