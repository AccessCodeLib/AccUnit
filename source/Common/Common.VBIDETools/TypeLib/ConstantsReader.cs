using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools.TypeLib
{
    public static class ConstantsReader
    {
        private static readonly Constants _constants = new Constants();

        public static void Clear()
        {
            _constants.Clear();
        }

        public static Constants Constants
        {
            get { return _constants; }
        }

        public static void AddConstants(VBProject vbProject)
        {
            _constants.Clear();
            foreach (Reference reference in vbProject.References)
            {
                try
                {
                    AddConstants(reference);
                }
                catch { }
            }
        }

        private static void AddConstants(Reference reference)
        {
            AddConstants(GetConstants(reference));
        }

        private static void AddConstants(Constants constants)
        {
            if (constants == null)
            {
                return;
            }
            foreach (var c in constants)
            {
                _constants.Add(c.Key, c.Value);
            }
        }

        private static Constants GetConstants(Reference reference)
        {
            try
            {
                var lib = new TypeLibInfo(reference.FullPath);
                return lib.Constants;
            }
            catch
            {
                return null;
            }

        }
    }
}
