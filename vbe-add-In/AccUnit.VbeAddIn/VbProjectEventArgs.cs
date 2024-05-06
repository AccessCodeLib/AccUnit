using Microsoft.Vbe.Interop;
using System;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class VbProjectEventArgs : EventArgs
    {
        private readonly VBProject _vbProject;

        public VbProjectEventArgs(VBProject vbProject)
        {
            _vbProject = vbProject;
        }

        public VBProject VBProject
        {
            get { return _vbProject; }
        }
    }
}
