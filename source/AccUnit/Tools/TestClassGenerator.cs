﻿using System.Collections.Generic;
using System.Linq;
using System.Net.Security;
using AccessCodeLib.AccUnit.Interop;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.Tools
{
    public class TestClassGenerator
    {
        private readonly VBProject _vbProject;

        public TestClassGenerator(VBProject vbproject)
        {
            _vbProject = vbproject;
        }

        public CodeModule NewTestClass(string CodeModuleToTest = null, bool CreateNewNameIfTestClassExists = true, bool GenerateTestMethodsFromCodeModuleToTest = false, string stateUnderTest = null, string expectedBehaviour = null)
        {
            var testClassName = GenerateTestClassName(CodeModuleToTest, CreateNewNameIfTestClassExists);
            string[] methodeNames = new string[] { };
            
            if (GenerateTestMethodsFromCodeModuleToTest)
            {
                methodeNames = GetMethodNamesFromCodeModule(CodeModuleToTest);
            }
            return InsertTestMethods(testClassName, methodeNames, stateUnderTest, expectedBehaviour);
        }

        private string[] GetMethodNamesFromCodeModule(string codeModuleToTest)
        {
            var codeModule = new CodeModuleContainer(_vbProject).TryGetCodeModule(codeModuleToTest);
            if (codeModule == null)
            {
                return new string[] { };
            }

            var codeModulueReader = new CodeModuleReader(codeModule);
            var members = codeModulueReader.Members;
            var publicMembers = members.FindAll(true).Select(member => member.Name).ToArray();

            return publicMembers;
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

        public CodeModule InsertTestMethods(string testClass, IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            var testCodeGenerator = new TestCodeGenerator();
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