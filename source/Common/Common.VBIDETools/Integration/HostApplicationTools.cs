using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System;

namespace AccessCodeLib.Common.VBIDETools.Integration
{
    public static class HostApplicationTools
    {
        const string AccessApplicationProgID = "Access.Application";

        public static object FindHostApplicationFromVBE(VBE vbe, out bool isAccessApplication)
        {
            using (new BlockLogger())
            {
                var progID = FindApplicationProgIdFromVBE(vbe);
                if (string.IsNullOrEmpty(progID))
                {
                    isAccessApplication = false;
                    return null;
                }
                isAccessApplication = progID.Equals(AccessApplicationProgID, StringComparison.CurrentCultureIgnoreCase);
                try
                {
                    return ComTools.GetOfficeComObjectByProgIDAndVBE(progID, vbe);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                    return null;
                }
            }
        }

        private static string FindApplicationProgIdFromVBE(VBE vbe)
        {
            using (new BlockLogger())
            {
                var activeVbProject = vbe.ActiveVBProject;
                if (activeVbProject != null)
                {
                    try
                    {
                        var progID = $"{activeVbProject.References.Item(2).Name}.Application";
                        // is this a stable statement?
                        Logger.Log($"Found ProgID in the active project's references: \"{progID}\".");
                        return progID;
                    }
                    catch (NullReferenceException ex)
                    {
                        Logger.Log(ex);
                        return FindApplicationProgIdFromVbeMenu(vbe);
                    }
                }
                else
                {
                    return FindApplicationProgIdFromVbeMenu(vbe);
                }
            }
        }

        private static string FindApplicationProgIdFromVbeMenu(VBE vbe)
        {
            using (new BlockLogger())
            {
                try
                {
                    var cbarControl = vbe.CommandBars["File"].FindControl(Id: 752);
                    var cbarCaption = cbarControl.Caption;
                    var appName = cbarCaption.Substring(cbarCaption.LastIndexOf(" ")).Trim();
                    var progID = $"{appName}.Application";
                    Logger.Log($"Found ProgID in VBE menu: \"{progID}\".");
                    return progID;
                }
                catch (NullReferenceException ex)
                {
                    Logger.Log(ex);
                    return null;
                }
            }
        }

        private static Type GetAccessApplicationTypeForComObject(object accessComObject)
        {
            using (new BlockLogger())
            {
                return ComTools.GetTypeForComObject(accessComObject, AccessApplicationProgID);
            }
        }

        public static OfficeApplicationHelper GetOfficeApplicationHelper(VBE vbe, ref object hostApplication)
        {
            using (new BlockLogger())
            {
                bool isAccessApplication;
                if (hostApplication is null)
                {
                    hostApplication = FindHostApplicationFromVBE(vbe, out isAccessApplication);
                }
                else
                {
                    isAccessApplication = GetAccessApplicationTypeForComObject(hostApplication) != null;
                }

                return hostApplication is null ? new VbeOnlyApplicatonHelper(vbe) : GetOfficeApplicationHelper(hostApplication, isAccessApplication);
            }
        }

        private static OfficeApplicationHelper GetOfficeApplicationHelper(object hostApplication, bool isAccessApplication)
        {
            using (new BlockLogger())
            {
                OfficeApplicationHelper applicationHelper;
                if (isAccessApplication)
                {
                    Logger.Log("Access application");
                    applicationHelper = new AccessApplicationHelper(hostApplication);
                }
                else
                {
                    applicationHelper = new OfficeApplicationHelper(hostApplication);
                }
                return applicationHelper;
            }
        }

    }
}