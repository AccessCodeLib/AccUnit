using System.Collections.Generic;
using AccessCodeLib.AccUnit.Configuration;
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

        public CodeModule InsertTestMethods(string testClass, IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            Configurator.CheckAccUnitVBAReferences(_vbProject);

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
