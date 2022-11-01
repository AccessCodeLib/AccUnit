using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools.Integration
{
    public static class ComTools
    {
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        private static IRunningObjectTable RunningObjectTable
        {
            get
            {
                IRunningObjectTable runningObjectTable;
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                {
                    return null;
                }
                return runningObjectTable;
            }
        }
        /// <summary>
        /// application object from VBE referenz and progID
        /// </summary>
        /// <param name="progID">example: Access.Application</param>
        /// <param name="vbe">object reference to VBE</param>
        /// <returns>Application</returns>
        /// <remarks>
        /// Microsoft Excel, PowerPoint, Word: enable option 'Trust access to the VBA project object model' (Options -> Trust Center -> Macro Settings)
        /// </remarks>
        public static object GetOfficeComObjectByProgIDAndVBE(string progID, VBE vbe)
        {
            using (new BlockLogger($"search: {progID}"))
            {
                object comInstance;

                // 1. Access.Application: use file name
                if (progID.Equals("Access.Application", StringComparison.CurrentCultureIgnoreCase) && vbe.VBProjects.Count == 1)
                {
                    var fileName = vbe.ActiveVBProject.FileName;
                    using (new BlockLogger($"Marshal.BindToMoniker({fileName})"))
                    {
                        comInstance = Marshal.BindToMoniker(fileName);
                        if (comInstance != null)
                        {
                            return comInstance;
                        }
                    }
                }

                // 2. attempt 'first' object in ROT
                using (new BlockLogger($"Marshal.GetActiveObject({progID})"))
                {
                    comInstance = GetOfficeComObjectByProgIDAndVBEfromFirstObjectInROT(progID, vbe);
                    if (comInstance != null)
                    {
                        return comInstance;
                    }
                }

                // 3. search in ROT
                using (new BlockLogger("GetOfficeComObjectByProgIDAndVBEfromROT"))
                {
                    return GetOfficeComObjectByProgIDAndVBEfromROT(progID, vbe);
                }
            }
        }

        private static object GetOfficeComObjectByProgIDAndVBEfromFirstObjectInROT(string progID, VBE vbe)
        {
            var comInstance = Marshal.GetActiveObject(progID);
            if (comInstance == null)
                return null;

            try
            {
                var appVbe = comInstance.GetType().InvokeMember("VBE", BindingFlags.GetProperty, null, comInstance, null);
                return vbe.Equals(appVbe) ? comInstance : null;
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                return null;
            }
        }

        private static object GetOfficeComObjectByProgIDAndVBEfromROT(string progID, VBE vbe)
        {
            using (new BlockLogger())
            {
                var runningObjectTable = RunningObjectTable;
                if (runningObjectTable == null)
                    return null;

                IEnumMoniker monikerList;
                runningObjectTable.EnumRunning(out monikerList);

                var monikerContainer = new IMoniker[1];
                var pointerFetchedMonikers = IntPtr.Zero;
                monikerList.Reset();
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    string displayName;
                    IBindCtx binder;
                    CreateBindCtx(0, out binder);

                    monikerContainer[0].GetDisplayName(binder, null, out displayName);
                    Logger.Log($"displayName: {displayName}");

                    object comInstance;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);
                    if (GetTypeForComObject(comInstance, progID) == null) continue;
                    object appVbe;
                    try
                    {
                        appVbe = comInstance.GetType().InvokeMember("VBE", BindingFlags.GetProperty, null, comInstance, null);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"VBE error => attempt: Application / Error: {ex.Message}");
                        try
                        {
                            comInstance = comInstance.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, comInstance, null);
                            // Microsoft Word & Co.: enable option 'Trust access to the VBA project object model'!
                            appVbe = comInstance.GetType().InvokeMember("VBE", BindingFlags.GetProperty, null, comInstance, null);
                            Logger.Log($"VBE: {comInstance.GetType()}");
                        }
                        catch (Exception ex2)
                        {
                            Logger.Log(ex2);
                            appVbe = comInstance = null;
                        }
                    }

                    if (appVbe == null || !vbe.Equals(appVbe)) continue;

                    Logger.Log($"found: {comInstance}");
                    return comInstance;
                }
                return null;
            }
        }

        /// <summary>
        /// code based on:
        /// http://fernandof.wordpress.com/2008/02/05/how-to-check-the-type-of-a-com-object-system__comobject-with-visual-c-net/
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="progID">ProgID for Type (example: Access.Application)</param>
        /// <returns></returns>
        public static Type GetTypeForComObject(object comObject, string progID)
        {
            using (new BlockLogger("comObject: " + comObject.GetType().Name))
            {
                var comType = Type.GetTypeFromProgID(progID);

                // get the com object and fetch its IUnknown
                var iunkwn = Marshal.GetIUnknownForObject(comObject);

                // enum all the types defined in the interop assembly
                var comAssembly = Assembly.GetAssembly(comType);
                var comTypes = comAssembly.GetTypes();

                using (new BlockLogger("foreach loop, comTypes.Count = " + comTypes.Length))
                {
                    // find the first implemented interop type
                    foreach (var currType in comTypes)
                    {
                        // get the iid of the current type
                        var iid = currType.GUID;
                        if (!currType.IsInterface || iid == Guid.Empty)
                        {
                            // com interop type must be an interface with valid iid
                            continue;
                        }

                        // query supportability of current interface on object
                        IntPtr ipointer;
                        Marshal.QueryInterface(iunkwn, ref iid, out ipointer);

                        if (ipointer != IntPtr.Zero)
                        {
                            // yeah, that’s the one we’re after
                            return currType;
                        }
                    }
                }
                // no implemented type found
                return null;
            }
        }

    }
}
