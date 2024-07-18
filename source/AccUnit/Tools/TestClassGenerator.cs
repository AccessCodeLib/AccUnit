using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit.Tools
{
    public class TestClassGenerator
    {
        private readonly VBProject _vbProject;
        private readonly ITestMethodBuilder _testMethodBuilder;

        public TestClassGenerator(VBProject vbproject, ITestMethodBuilder testMethodBuilder)
        {
            _vbProject = vbproject;
            _testMethodBuilder = testMethodBuilder;
        }

        public CodeModule NewTestClass(string CodeModuleToTest = null,
                                        bool CreateNewNameIfTestClassExists = true,
                                        bool GenerateTestMethodsFromCodeModuleToTest = false,
                                        string stateUnderTest = null, string expectedBehaviour = null)
        {
            var testClassName = GenerateTestClassName(CodeModuleToTest, CreateNewNameIfTestClassExists);
            IEnumerable<CodeModuleMember> methods = new TestCodeModuleMember[] { };

            if (GenerateTestMethodsFromCodeModuleToTest)
            {
                methods = GetTestCodeModuleMemberFromCodeModule(CodeModuleToTest, stateUnderTest, expectedBehaviour);
            }
            return InsertTestMethods(testClassName, methods);
        }

        private IEnumerable<CodeModuleMember> GetTestCodeModuleMemberFromCodeModule(
                string codeModuleToTest, string stateUnderTest = null, string expectedBehaviour = null,
                IEnumerable<string> methodsUnderTestFilter = null)
        {
            var codeModule = new CodeModuleContainer(_vbProject).TryGetCodeModule(codeModuleToTest);
            if (codeModule is null)
            {
                return new TestCodeModuleMember[] { };
            }

            var codeModulueReader = new CodeModuleReader(codeModule);
            var members = codeModulueReader.Members;
            var publicMembers = members.FindAll(true);

            if (methodsUnderTestFilter != null)
            {
                publicMembers = publicMembers.Where(m => methodsUnderTestFilter.Contains(m.Name)).ToList();
            }

            var testCodeModuleMembers = new List<CodeModuleMember>();
            foreach (var member in publicMembers)
            {
                testCodeModuleMembers.Add(new TestCodeModuleMember(member, stateUnderTest, expectedBehaviour));
            }

            return testCodeModuleMembers;
        }

        protected bool CodeModuleExists(string testClass)
        {
            var modules = new CodeModuleContainer(_vbProject);
            return modules.TryGetCodeModule(testClass) != null;
        }

        protected string GenerateTestClassName(string CodeModuleToTest, bool CreateNewNameIfTestClassExists = false)
        {
            // e. g. TestClassNameFormat = %ModuleUnderTest%Tests
            var testClassName = Properties.Settings.Default.TestClassNameFormat.Replace("%ModuleUnderTest%", CodeModuleToTest);
            if (CreateNewNameIfTestClassExists)
            {
                if (CodeModuleExists(testClassName))
                {
                    int i = 0;
                    string testClassNameWithNumber;
                    do
                    {
                        i++;
                        testClassNameWithNumber = testClassName + i.ToString();
                    } while (CodeModuleExists(testClassNameWithNumber));
                    testClassName = testClassNameWithNumber;
                }
            }
            return testClassName;
        }

        public CodeModule InsertTestMethods(string testClass, IEnumerable<CodeModuleMember> testMethods)
        {
            var testCodeGenerator = new TestCodeGenerator(_testMethodBuilder);
            testCodeGenerator.Add(testMethods);

            var modules = new CodeModuleContainer(_vbProject);
            var testModule = modules.TryGetCodeModule(testClass);
            var testModuleExists = testModule != null;
            var sourcecode = testCodeGenerator.GenerateSourceCode(includeHeader: !testModuleExists);

            if (testModuleExists)
            {
                testModule.InsertLines(testModule.CountOfLines + 1, sourcecode);
                return testModule;
            }
            return modules.Generator.Add(vbext_ComponentType.vbext_ct_ClassModule, testClass, sourcecode, true);
        }

        public CodeModule InsertTestMethods(string codeModuleToTest, string testClass, IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            if (codeModuleToTest == null        )
                return InsertTestMethods(testClass, methodsUnderTest, stateUnderTest, expectedBehaviour);

            var methods = GetTestCodeModuleMemberFromCodeModule(codeModuleToTest, stateUnderTest, expectedBehaviour, methodsUnderTest);
            return InsertTestMethods(testClass, methods);
        }

        public CodeModule InsertTestMethods(string testClass, IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            var testCodeGenerator = new TestCodeGenerator(_testMethodBuilder);
            testCodeGenerator.Add(methodsUnderTest, stateUnderTest, expectedBehaviour);

            var modules = new CodeModuleContainer(_vbProject);
            var testModule = modules.TryGetCodeModule(testClass);
            var testModuleExists = testModule != null;
            var sourcecode = testCodeGenerator.GenerateSourceCode(includeHeader: !testModuleExists);

            if (testModuleExists)
            {
                testModule.InsertLines(testModule.CountOfLines + 1, sourcecode);
                return testModule;
            }
            return modules.Generator.Add(vbext_ComponentType.vbext_ct_ClassModule, testClass, sourcecode, true);
        }
    }
}
