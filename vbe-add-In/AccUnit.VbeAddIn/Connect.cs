using System;
using System.Runtime.InteropServices;
using AccessCodeLib.Common.Tools.Logging;

namespace AccessCodeLib.AccUnit.VbeAddIn
{

    #region Read me for Add-in installation and setup information.
    // When run, the Add-in wizard prepared the registry for the Add-in.
    // At a later time, if the Add-in becomes unavailable for reasons such as:
    //   1) You moved this project to a computer other than which is was originally created on.
    //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
    //   3) Registry corruption.
    // you will need to re-register the Add-in by building the MyAddin1Setup project, 
    // right click the project in the Solution Explorer, then choose install.
    #endregion

    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [ComVisible(true)]
    [Guid("F15F18C3-CA43-421E-9585-6A04F51C5786")]
    [ProgId(ComRegistration.ComProgId)]
    public class Connect : Object, Extensibility.IDTExtensibility2, IDisposable
    {
        #region IDTExtensibility2 implementation

        Microsoft.Vbe.Interop.AddIn _addInInstance;
        private AddInManager AddInManager { get; set; }

        /// <summary>
        ///      Implements the OnConnection method of the IDTExtensibility2 interface.
        ///      Receives notification that the Add-in is being loaded.
        /// </summary>
        /// <param name='application'>
        ///      Root object of the host application.
        /// </param>
        /// <param name='connectMode'>
        ///      Describes how the Add-in is being loaded.
        /// </param>
        /// <param name='addInInstance'>
        ///      Object representing this Add-in.
        /// </param>
        /// <param name="custom"></param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInstance, ref Array custom)
        {
            using (new BlockLogger())
            {
                try
                {
                    _addInInstance = (Microsoft.Vbe.Interop.AddIn)addInInstance;
                    AddInManager = new AddInManager(_addInInstance);
                    AddInManager.Init();

                    _addInInstance.Object = AddInManager.Bridge;
                }
                catch (Exception ex)
                {
                    UITools.ShowException(ex);
                }
            }
        }

        /// <summary>
        ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
        ///     Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param name='disconnectMode'>
        ///      Describes how the Add-in is being unloaded.
        /// </param>
        /// <param name='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref Array custom)
        {
            using (new BlockLogger())
            {
                try
                {
                    Dispose();
                }
                catch(Exception xcp)
                {
                    Logger.Log(xcp);
                }
            }
        }

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param name='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref Array custom)
        {
            Logger.Log();
        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param name='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
            Logger.Log();
        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param name='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                using (new BlockLogger())
                {
                    try
                    {
                        _addInInstance.Object = null;
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(ex);
                    }
                }
            }
            catch { /* ignore */}
        }

        #endregion

        #region COM register/unregister support
        [ComRegisterFunction]
        public static void RegisterClass(Type t)
        {
            ComRegistration.ComRegisterClass(t);
        }

        [ComUnregisterFunction]
        public static void UnregisterClass(Type t)
        {
            ComRegistration.ComUnregisterClass(t);
        }

        #endregion

        #region IDisposable Support

        public delegate void DisposeEventHandler(object sender);
        public event DisposeEventHandler Disposed;

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            using (new BlockLogger())
            {
                if (_disposed)
                {
                    Logger.Log("_disposed == true => Exit Dispose()");
                    return;
                }

                if (disposing)
                {
                    Logger.Log("disposing == true");
                    if (AddInManager != null)
                    {
                        Logger.Log("Start AddInManager.Dispose ...");
                        AddInManager.Dispose();
                        AddInManager = null;
                    }
                }

                _addInInstance = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                _disposed = true;

                if (Disposed != null)
                {
                    Disposed(this);
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~Connect()
        {
            Logger.Log("~Connect");
            Dispose(false);
            Logger.Log("~Connect completed");
        }

        #endregion
    }
}
