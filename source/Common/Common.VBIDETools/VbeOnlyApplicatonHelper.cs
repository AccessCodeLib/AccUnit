using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools
{
    public class VbeOnlyApplicatonHelper : OfficeApplicationHelper
    {
        private readonly VBE _vbe;

        public VbeOnlyApplicatonHelper(VBE vbe)
            : base(vbe)
        {
            _vbe = vbe;
        }

        public override VBE VBE
        {
            get { return _vbe; }
        }

        protected override VBProject GetCheckedVbProject()
        {
            using (new BlockLogger())
            {
                return _vbe.ActiveVBProject;
            }
        }
    }
}