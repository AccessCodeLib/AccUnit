using Microsoft.Vbe.Interop;
using System.Collections.Generic;

namespace AccessCodeLib.Common.VBIDETools
{
    public class CodeModuleMemberList : List<CodeModuleMember>
    {
        public List<CodeModuleMember> FindAll(bool isPublic)
        {
            return FindAll(
                member => member.IsPublic == isPublic
                );
        }

        public List<CodeModuleMember> FindAll(bool isPublic, vbext_ProcKind procKind)
        {
            return FindAll(
                member => member.IsPublic == isPublic && member.ProcKind == procKind
                );
        }

    }
}
