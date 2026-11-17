using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace AccessCodeLib.Common.VBIDETools.Integration
{
    public static class ComTools
    {
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        [DllImport("ole32.dll")]
        private static extern int CreateFileMoniker([MarshalAs(UnmanagedType.LPWStr)] string lpszPathName, out IMoniker ppmk);

        private static IRunningObjectTable RunningObjectTable
        {
            get
            {
                if (GetRunningObjectTable(0, out IRunningObjectTable runningObjectTable) != 0 || runningObjectTable is null)
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
                    using (new BlockLogger($"GetComObjectFromROTByFileName({fileName})"))
                    {
                        comInstance = GetComObjectFromROTByFileName(fileName);
                        if (comInstance != null)
                        {
                            if (BelongsToVbe(comInstance, vbe))
                            {
                                return comInstance;
                            }
                            ReleaseComObject(comInstance);
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

        /// <summary>
        /// Looks up a file moniker in the running object table.
        /// </summary>
        /// <remarks>
        /// Unlike Marshal.BindToMoniker, IRunningObjectTable.GetObject never activates the object,
        /// so no new instance of the host application is started if the file is not open.
        /// </remarks>
        private static object GetComObjectFromROTByFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return null;

            var runningObjectTable = RunningObjectTable;
            if (runningObjectTable is null)
                return null;

            IMoniker fileMoniker = null;
            try
            {
                if (CreateFileMoniker(fileName, out fileMoniker) != 0 || fileMoniker is null)
                    return null;

                return runningObjectTable.GetObject(fileMoniker, out object comInstance) == 0 ? comInstance : null;
            }
            finally
            {
                ReleaseComObject(fileMoniker);
                ReleaseComObject(runningObjectTable);
            }
        }

        private static object GetOfficeComObjectByProgIDAndVBEfromFirstObjectInROT(string progID, VBE vbe)
        {
            var comInstance = Marshal.GetActiveObject(progID);
            if (comInstance is null)
                return null;

            if (BelongsToVbe(comInstance, vbe))
                return comInstance;

            ReleaseComObject(comInstance);
            return null;
        }

        private static object GetOfficeComObjectByProgIDAndVBEfromROT(string progID, VBE vbe)
        {
            using (new BlockLogger())
            {
                var runningObjectTable = RunningObjectTable;
                if (runningObjectTable is null)
                    return null;

                IEnumMoniker monikerList = null;
                try
                {
                    runningObjectTable.EnumRunning(out monikerList);

                    var monikerContainer = new IMoniker[1];
                    var pointerFetchedMonikers = IntPtr.Zero;
                    monikerList.Reset();
                    while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                    {
                        var moniker = monikerContainer[0];
                        monikerContainer[0] = null;
                        try
                        {
                            var comInstance = GetComObjectFromROTIfOfficeApplicationOfVbe(runningObjectTable, moniker, progID, vbe);
                            if (comInstance != null)
                            {
                                Logger.Log($"found: {comInstance}");
                                return comInstance;
                            }
                        }
                        finally
                        {
                            ReleaseComObject(moniker);
                        }
                    }
                    return null;
                }
                finally
                {
                    ReleaseComObject(monikerList);
                    ReleaseComObject(runningObjectTable);
                }
            }
        }

        private static object GetComObjectFromROTIfOfficeApplicationOfVbe(IRunningObjectTable runningObjectTable, IMoniker moniker, string progID, VBE vbe)
        {
            IBindCtx binder = null;
            try
            {
                CreateBindCtx(0, out binder);
                moniker.GetDisplayName(binder, null, out string displayName);
                Logger.Log($"displayName: {displayName}");
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
            finally
            {
                ReleaseComObject(binder);
            }

            if (runningObjectTable.GetObject(moniker, out object comInstance) != 0 || comInstance is null)
                return null;

            if (GetTypeForComObject(comInstance, progID) is null)
            {
                ReleaseComObject(comInstance);
                return null;
            }

            if (BelongsToVbe(comInstance, vbe))
                return comInstance;

            // Microsoft Word & Co.: enable option 'Trust access to the VBA project object model'!
            Logger.Log("VBE does not match or is not accessible => attempt: Application");
            object application = null;
            try
            {
                application = comInstance.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, comInstance, null);
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
            finally
            {
                // each marshalled reference increments the RCW count, so this is balanced
                // even if Application returns the same RCW
                ReleaseComObject(comInstance);
            }

            if (application is null)
                return null;

            if (BelongsToVbe(application, vbe))
            {
                Logger.Log($"VBE: {application.GetType()}");
                return application;
            }

            ReleaseComObject(application);
            return null;
        }

        /// <summary>
        /// Checks whether the VBE property of the COM object is the given VBE.
        /// </summary>
        private static bool BelongsToVbe(object comInstance, VBE vbe)
        {
            object appVbe = null;
            try
            {
                appVbe = comInstance.GetType().InvokeMember("VBE", BindingFlags.GetProperty, null, comInstance, null);
                return appVbe != null && vbe.Equals(appVbe);
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                return false;
            }
            finally
            {
                // the matching VBE is the RCW shared with the add-in => do not release it
                if (!ReferenceEquals(appVbe, vbe))
                {
                    ReleaseComObject(appVbe);
                }
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject is null || !Marshal.IsComObject(comObject))
                return;

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
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
                try
                {
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
                            Marshal.QueryInterface(iunkwn, ref iid, out IntPtr ipointer);

                            if (ipointer != IntPtr.Zero)
                            {
                                Marshal.Release(ipointer);
                                // yeah, that’s the one we’re after
                                return currType;
                            }
                        }
                    }
                    // no implemented type found
                    return null;
                }
                finally
                {
                    Marshal.Release(iunkwn);
                }
            }
        }

    }
}
