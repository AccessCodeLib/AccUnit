using System;
using System.Linq;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.Configuration
{
    public class AccUnitVBAReference
    {
        public AccUnitVBAReference(string guid, int major, int minor)
        {
            Guid = guid;
            Major = major;
            Minor = minor;
        }

        public string Guid { get; private set; }
        public int Major { get; private set; }
        public int Minor { get; private set; }

        public void EnsureReferenceExistsInVbProject(_VBProject vbProject)
        {
            var referenceToSelf = GetReferenceToSelf(vbProject);

            var referenceIsOk = false;
            if (referenceToSelf != null)
            {
                if (IsReferenceCurrentOrNewer(referenceToSelf))
                {
                    referenceIsOk = true;
                }
                else // remove reference to add reference with newer major/minor
                {
                    RemoveReferenceFromVbProject(vbProject, referenceToSelf);
                }
            }

            if (!referenceIsOk)
            {
                AddReferenceToSelf(vbProject);
            }
        }

        public void RemoveFromVbProject(_VBProject vbProject)
        {
            foreach (var reference in
                vbProject.References.Cast<Reference>().Where(IsReferenceToSelf))
            {
                RemoveReferenceFromVbProject(vbProject, reference);
                break;
            }
        }

        private Reference GetReferenceToSelf(_VBProject vbProject)
        {
            return vbProject.References.Cast<Reference>().FirstOrDefault(IsReferenceToSelf);
        }

        private static void RemoveReferenceFromVbProject(_VBProject vbProject, Reference reference)
        {
            vbProject.References.Remove(reference);
        }

        private bool IsReferenceToSelf(Reference reference)
        {
            return reference.Guid.Equals(Guid, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsReferenceCurrentOrNewer(Reference reference)
        {
            return reference.Major >= Major && reference.Minor >= Minor;
        }

        private void AddReferenceToSelf(_VBProject vbProject)
        {
            vbProject.References.AddFromGuid(Guid, Major, Minor);
        }
    }
}