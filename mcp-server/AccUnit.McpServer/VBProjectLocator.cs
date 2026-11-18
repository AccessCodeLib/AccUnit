using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Access = Microsoft.Office.Interop.Access;

namespace AccessCodeLib.AccUnit.McpServer
{
    internal static class VBProjectLocator
    {
        /// <summary>
        /// VBE.ActiveVBProject can point to a loaded library (e.g. a protected Access wizard) after code ran,
        /// so the project is found by the file name of the open database.
        /// </summary>
        public static VBProject GetCurrentVBProject(Access.Application application)
        {
            var databasePath = GetCurrentProjectFullName(application);
            var vbe = (VBE)application.VBE;
            foreach (VBProject vbProject in vbe.VBProjects)
            {
                if (string.Equals(GetFileNameOrNull(vbProject), databasePath, StringComparison.OrdinalIgnoreCase))
                    return vbProject;
            }
            throw new DatabaseSessionException("The VBA project of the open database was not found: " + databasePath);
        }

        private static string GetCurrentProjectFullName(Access.Application application)
        {
            var currentProject = application.CurrentProject;
            try
            {
                return currentProject.FullName;
            }
            finally
            {
                Marshal.ReleaseComObject(currentProject);
            }
        }

        private static string GetFileNameOrNull(VBProject vbProject)
        {
            try
            {
                return vbProject.FileName;
            }
            catch (COMException)
            {
                // Protected or not yet saved projects do not expose their file name.
                return null;
            }
        }
    }
}
