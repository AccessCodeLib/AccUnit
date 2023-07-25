using Microsoft.Vbe.Interop;
using System.Reflection;
using System.Runtime.InteropServices;
using static System.String;

namespace AccessCodeLib.AccUnit.Configuration
{
    public class AccUnitVBAReferences
    {
        private AccUnitVBAReference _accUnitReference;
        public AccUnitVBAReference AccUnitReference
        {
            get { return _accUnitReference ?? (_accUnitReference = GetAccUnitReference()); }
        }

        private static AccUnitVBAReference GetAccUnitReference()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var guidString = $"{{{GetAssemblyGuidString(assembly).ToUpper()}}}";
            var majorVersion = Assembly.GetExecutingAssembly().GetName().Version.Major;
            var minorVersion = Assembly.GetExecutingAssembly().GetName().Version.Minor;

            return new AccUnitVBAReference(guidString, majorVersion, minorVersion);
        }

        private static string GetAssemblyGuidString(ICustomAttributeProvider assembly)
        {
            var guidAttributes = assembly.GetCustomAttributes(typeof(GuidAttribute), false);
            return guidAttributes.Length <= 0 ? Empty : ((GuidAttribute)guidAttributes[0]).Value;
        }

        public void EnsureReferencesExistIn(VBProject vbProject)
        {
            AccUnitReference.EnsureReferenceExistsInVbProject(vbProject);
        }

        public void RemoveReferencesFrom(VBProject vbProject)
        {
            AccUnitReference.RemoveFromVbProject(vbProject);
        }
    }
}