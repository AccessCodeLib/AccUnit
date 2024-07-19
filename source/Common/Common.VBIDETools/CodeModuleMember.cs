using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools
{
    public class CodeModuleMember
    {
        public CodeModuleMember(string name, vbext_ProcKind procKind, bool isPublic, 
                                string declarationString = "", 
                                string codeModuleName = null,
                                string procedureCode = null)
        {
            using (new BlockLogger())
            {
                Name = name;
                ProcKind = procKind;
                IsPublic = isPublic;
                DeclarationString = declarationString;
                CodeModuleName = codeModuleName;
                ProcedureCode = procedureCode;
            }
        }

        public string Name { get; }
        public vbext_ProcKind ProcKind { get; }
        public bool IsPublic { get; }
        public string DeclarationString { get; }
        public string ProcedureCode { get; }
        public string CodeModuleName { get; set; }
    }
}
