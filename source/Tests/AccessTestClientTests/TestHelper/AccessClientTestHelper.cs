using AccessCodeLib.Common.TestHelpers.AccessRelated;
using Microsoft.Office.Interop.Access;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    internal static class AccessClientTestHelper
    {
        private static int staticClientCounter;

        public static CodeModule CreateTestCodeModule(AccessTestHelper accessTestHelper, string name, vbext_ComponentType type, string source)
        {
            var vbcCol = accessTestHelper.ActiveVBProject.VBComponents;
            var vbc = vbcCol.Add(type);
            vbc.Name = name;
            if (type == vbext_ComponentType.vbext_ct_ClassModule)
                vbc.Properties.Item("Instancing").Value = 2; // 2 = Public, damit aus Test aufrufbar
            accessTestHelper.Application.RunCommand(AcCommand.acCmdCompileAndSaveAllModules);
            var codeModule = vbc.CodeModule;
            codeModule.Name = name;
            codeModule.DeleteLines(1, codeModule.CountOfLines);
            codeModule.InsertLines(1, source);
            accessTestHelper.Application.RunCommand(AcCommand.acCmdCompileAndSaveAllModules);
            return codeModule;
        }

        public static AccessTestHelper NewAccessTestHelper()
        {
            var testHelper = new AccessTestHelper(@"C:\test\Test_" + staticClientCounter++.ToString() + ".accdb");
            testHelper.Application.Visible = true;

            return testHelper;
        }
    }
}
