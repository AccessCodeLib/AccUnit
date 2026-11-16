using System;
using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.AccUnit.Tools;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Creates the draft of a test class for a module of the open database with the AccUnit test code generator
    /// (as the VBE add-in does). Reads the VBA project only; nothing is written to the database.
    /// </summary>
    public sealed class TestClassDraftService
    {
        private readonly AccessSession _session;

        public TestClassDraftService(AccessSession session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        public TestClassDraft CreateDraft(string moduleName, IList<string> methodNames)
        {
            if (string.IsNullOrWhiteSpace(moduleName))
                throw new DatabaseSessionException("The name of the module under test is missing.");

            return _session.Execute(application =>
                CreateDraftOnStaThread(VBProjectLocator.GetCurrentVBProject(application), moduleName, methodNames));
        }

        private static TestClassDraft CreateDraftOnStaThread(VBProject vbProject, string moduleName, IList<string> methodNames)
        {
            using (var modules = new CodeModuleContainer(vbProject))
            {
                var codeModule = modules.TryGetCodeModule(moduleName);
                if (codeModule == null)
                    throw new DatabaseSessionException("Module not found: " + moduleName);

                return CreateDraft(vbProject, new CodeModuleReader(codeModule), moduleName, methodNames);
            }
        }

        private static TestClassDraft CreateDraft(VBProject vbProject, CodeModuleReader reader, string moduleName, IList<string> methodNames)
        {
            var members = SelectMembers(reader.Members.FindAll(true), methodNames);

            var generator = new TestCodeGenerator();
            generator.Add(members);

            return new TestClassDraft
            {
                TestClassName = new TestClassNameGenerator(vbProject).GetTestClassName(moduleName),
                ModuleUnderTest = moduleName,
                SourceCode = generator.GenerateSourceCode(),
                MethodsUnderTest = members.Select(member => new MethodUnderTest
                {
                    Name = member.Name,
                    Kind = GetKind(member.ProcKind),
                    Declaration = member.DeclarationString,
                    Comment = ProcedureCommentFormatter.Format(reader.GetProcedureHeader(member.Name, member.ProcKind))
                }).ToList()
            };
        }

        private static List<CodeModuleMember> SelectMembers(List<CodeModuleMember> publicMembers, IList<string> methodNames)
        {
            if (methodNames == null || methodNames.Count == 0)
                return publicMembers;

            var unknown = methodNames.Where(name => !publicMembers.Any(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase))).ToList();
            if (unknown.Count > 0)
                throw new DatabaseSessionException("No public method with this name: " + string.Join(", ", unknown));

            return publicMembers.Where(m => methodNames.Contains(m.Name, StringComparer.OrdinalIgnoreCase)).ToList();
        }

        private static string GetKind(vbext_ProcKind procKind)
        {
            switch (procKind)
            {
                case vbext_ProcKind.vbext_pk_Get: return "Property Get";
                case vbext_ProcKind.vbext_pk_Let: return "Property Let";
                case vbext_ProcKind.vbext_pk_Set: return "Property Set";
                default: return "Procedure";
            }
        }

        /// <summary>Uses the naming of the AccUnit test class generator (setting TestClassNameFormat, e.g. "...Tests").</summary>
        private sealed class TestClassNameGenerator : TestClassGenerator
        {
            public TestClassNameGenerator(VBProject vbProject) : base(vbProject)
            {
            }

            public string GetTestClassName(string moduleUnderTest)
            {
                return GenerateTestClassName(moduleUnderTest, true);
            }
        }
    }
}
