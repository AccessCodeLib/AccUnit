using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools
{
    public class VBProjectOnlyApplicatonHelper : OfficeApplicationHelper
    {
        private readonly VBProject _vbProject;

        public VBProjectOnlyApplicatonHelper(VBProject vbProject)
            : base(vbProject)
        {
            _vbProject = vbProject;
        }

        public override VBE VBE
        {
            get { return _vbProject.VBE; }
        }

        protected override VBProject GetCheckedVbProject()
        {
            using (new BlockLogger())
            {
                return _vbProject;
            }
        }
    }
}