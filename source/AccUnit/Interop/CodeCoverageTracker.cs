using Microsoft.Vbe.Interop;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("ED31BE77-E17D-42FA-95F9-5280798B22CD")]
    public interface ICodeCoverageTracker
    {
        ICodeCoverageTracker Add(string CodeModuleName);
        void Track(string CodeModulName, string ProcedureName, int LineNo);
        string GetReport();
    }

    [ComVisible(true)]
    [Guid("6DD0F6D6-D2E0-4AB9-8F51-8CA4011EFD89")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".CodeCoverageTracker")]
    public class CodeCoverageTracker : CodeCoverage.CodeCoverageTracker, ICodeCoverageTracker
    {
        public CodeCoverageTracker(VBProject vbProject = null) : base(vbProject)
        {
        }

        public new ICodeCoverageTracker Add(string CodeModuleName)
        {
            base.Add(CodeModuleName);
            return this;
        }
    }

}
