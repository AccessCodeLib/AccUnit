using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AccessCodeLib.AccUnit.Common;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit
{
    public interface ITestClassReader
    {
        TestClassList GetTestClasses(bool initMembers = false);
        IEnumerable<CodeModuleInfo> GetTestComponents();
        TestClassMemberList GetTestClassMemberList(string classname);
        TestClassMemberInfo GetTestClassMemberInfo(string classname, string membername);
    }

    public class TestClassReader : ITestClassReader
    {
        public TestClassReader(_VBProject vbProject)
        {
            if (vbProject == null)
                throw new ArgumentNullException(nameof(vbProject));
            VbProject = vbProject;
        }

        private _VBProject VbProject { get; }

        public TestClassList GetTestClasses(bool initMembers = false)
        {
            using (new BlockLogger())
            {
                var vbComponents = VbProject.VBComponents;
                return new TestClassList(from VBComponent vbc in vbComponents
                                         where vbc.Type == vbext_ComponentType.vbext_ct_ClassModule
                                         where IsTestClassCodeModul(vbc.CodeModule)
                                         select GetTestClassInfo(vbc, initMembers));
            }
        }

        public IEnumerable<CodeModuleInfo> GetTestComponents()
        {
            using (new BlockLogger())
            {
                var vbComponents = VbProject.VBComponents;
                return from VBComponent vbc in vbComponents
                       where IsTestRelatedCodeModule(vbc.CodeModule)
                       select new CodeModuleInfo(vbc.Name, vbc.Type);
            }
        }

        private static TestClassInfo GetTestClassInfo(_VBComponent vbc, bool initMembers)
        {
            using (new BlockLogger(vbc.CodeModule.Name))
            {
                var reader = new CodeModuleReader(vbc.CodeModule);
                var members = (initMembers ? GetTestClassMembers(reader) : null);
                return new TestClassInfo(vbc.Name, reader.SourceCode, members);
            }
        }

        private static TestClassMemberList GetTestClassMembers(CodeModuleReader reader)
        {
            var list = new TestClassMemberList();
            list.AddRange(
                    from CodeModuleMember member in reader.Members.FindAll(true, vbext_ProcKind.vbext_pk_Proc)
                    where IsSetupOrTeardown(member) == false
                    select GetTestClassMemberInfo(member, reader)
            );
            return list;
        }

        private static bool IsSetupOrTeardown(CodeModuleMember member)
        {
            switch (member.Name.ToLower())
            {
                case "fixturesetup":
                case "setup":
                case "teardown":
                case "fixtureteardown":
                    return true;
                default:
                    return false;
            }
        }

        public TestClassMemberList GetTestClassMemberList(string classname)
        {
            using (new BlockLogger())
            {
                var codeModule = VbProject.VBComponents.Item(classname).CodeModule;
                return GetTestClassMembers(new CodeModuleReader(codeModule));
            }
        }

        public TestClassMemberInfo GetTestClassMemberInfo(string classname, string membername)
        {
            using (new BlockLogger())
            {
                var reader = new CodeModuleReader(VbProject.VBComponents.Item(classname).CodeModule);
                var memberInfo = new TestClassMemberInfo(membername, reader.GetProcedureHeader(membername));

                var rowGenerator = new TestRowGenerator();
                rowGenerator.ActiveVBProject = (VBProject)VbProject;
                rowGenerator.TestName = classname;
                var testRows = rowGenerator.GetTestRows(membername);

                memberInfo.TestRows.AddRange(testRows);
                
                return memberInfo;
            }
        }

        private static TestClassMemberInfo GetTestClassMemberInfo(CodeModuleMember member, CodeModuleReader reader)
        {
            return new TestClassMemberInfo(member.Name, reader.GetProcedureHeader(member.Name, member.ProcKind));
        }

        public static bool IsTestClassCodeModul(_CodeModule codeModule)
        {
            return CodeModuleHeaderMatchesRegex(codeModule, @"^\s*'\s*AccUnit:TestClass\s*$");
        }

        private static int GetLinesBeforeFirstProcBody(_CodeModule codeModule)
        {
            vbext_ProcKind procKind;
            //Logger.Log(string.Format("Module: {0}", codeModule.Name));
            //Logger.Log(string.Format("Module: {0}\nCountOfLines: {1}", codeModule.Name, codeModule.CountOfLines));
            //Logger.Log(string.Format("Module: {0}\nCountOfDeclarationLines: {1}", codeModule.Name, codeModule.CountOfDeclarationLines));
// ReSharper disable UseIndexedProperty 
            var firstProc = codeModule.get_ProcOfLine(codeModule.CountOfDeclarationLines + 1, out procKind);
// ReSharper restore UseIndexedProperty
            return string.IsNullOrEmpty(firstProc) ? codeModule.CountOfLines : codeModule.ProcBodyLine[firstProc, procKind];
        }

        private static bool IsTestRelatedCodeModule(_CodeModule codeModule)
        {
            return CodeModuleHeaderMatchesRegex(codeModule, @"^\s*'\s*AccUnit:(TestClass|TestRelated)\s*$");
        }

        private static bool CodeModuleHeaderMatchesRegex(_CodeModule codeModule, string regexString)
        {
            var linesToCheck = GetLinesBeforeFirstProcBody(codeModule);
            var checkString = codeModule.Lines[1, linesToCheck];
            var regex = new Regex(regexString, RegexOptions.CultureInvariant | RegexOptions.Compiled |
                                               RegexOptions.Multiline | RegexOptions.IgnoreCase);
            return !string.IsNullOrEmpty(checkString) && regex.IsMatch(checkString);
        }

        public static List<System.IO.FileInfo> GetTestFilesFromDirectory(string path, string fileNameSeachPattern = "*")
        {
            var di = new System.IO.DirectoryInfo(path);
            var list = di.GetFiles(fileNameSeachPattern + ".cls").ToList();
            list.AddRange(di.GetFiles(fileNameSeachPattern + ".bas").ToList());
            list.AddRange(di.GetFiles(fileNameSeachPattern + ".acf").ToList());
            list.AddRange(di.GetFiles(fileNameSeachPattern + ".acr").ToList());
            return list;
        }

    }
}
