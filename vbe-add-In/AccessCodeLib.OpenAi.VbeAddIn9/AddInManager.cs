using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.OpenAI.VbeAddIn
{
    internal class AddInManager : IDisposable
    {
        private AddIn _addIn;
        
        public AddInManager(AddIn addIn)
        {
            _addIn = addIn;
        }

        public void Init()
        {
            /*
                if (!ApplicationSupportsAddIn)
                    return;

                try
                {
                    
                }
                catch (Exception ex)
                {
                    //UITools.ShowException(ex);
                }
            */
        }
        /*
        static void OnShowUIMessage(object sender, MessageEventArgs e)
        {
            UITools.ShowMessage(e.Message, e.Buttons, e.Icon, e.DefaultButton);
            e.MessageDisplayed = true;
        }
        */
        //private bool ApplicationSupportsAddIn => !(_officeApplicationHelper is VbeOnlyApplicatonHelper);

        
        #region ad Bridge

        private AddInManagerBridge _addInManagerBridge;

        public AddInManagerBridge Bridge
        {
            get
            {
                if (_addInManagerBridge == null)
                {
                    CreateAddInManagerBridge();
                }
                return _addInManagerBridge;
            }
        }

        private void CreateAddInManagerBridge()
        {
            _addInManagerBridge = new AddInManagerBridge();
            _addInManagerBridge.TestCodeBuilderFactoryRequest += AddInBridgeTestCodeBuilderFactoryRequest;
        }

        private void AddInBridgeTestCodeBuilderFactoryRequest(out ITestCodeBuilderFactory testCodeBuilderFactory)
        {
            testCodeBuilderFactory = new TestCodeBuilderFactory(new OpenAiService(new CredentialManager()));
        }

        void AddInBridgeHostApplicationInitialized(object application)
        {
            //InitOfficeApplicationHelper(application);
        }

        /*
        private void InitOfficeApplicationHelper(object hostApplication = null)
        {
                // Note: if load RubberDuck, an instance of Access stay in memory after close => HostApplicationTools.GetOfficeApplicationHelper(..., ..., true);
                _officeApplicationHelper = HostApplicationTools.GetOfficeApplicationHelper(VBE, ref hostApplication, true);
                _vbeIntegrationManager.OfficeApplicationHelper = _officeApplicationHelper;
                _testSuiteManager.OfficeApplicationHelper = _officeApplicationHelper;
        }
        */

        #endregion


        public static string FriendlyName => $"AccessCodeLib.OpenAI {FileVersion}";

        public static string FileVersion
        {
            get
            {
                var version = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                return version.FileVersion;
            }
        }

        public static string Copyright
        {
            get
            {
                var version = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                return version.LegalCopyright;
            }
        }

        private VBE VBE => AddIn.VBE;

        private AddIn AddIn => _addIn;


        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
                if (_disposed) return;

                if (disposing)
                {
                    try
                    {
                        DisposeManagedResources();
                    }
                    catch (Exception ex)
                    {
                        //Logger.Log(ex);
                    }
                }

                try
                {
                    DisposeUnmanagedResources();
                }
                catch (Exception ex)
                {
                    //
                }

                _disposed = true;
        }

        private void DisposeUnmanagedResources()
        {
            if (_addIn != null)
            {
                Marshal.ReleaseComObject(_addIn);
                _addIn = null;
            }
        }

        private void DisposeManagedResources()
        {
                DisposeAddInManagerBridge();
        }

        private void DisposeAddInManagerBridge()
        {
                if (_addInManagerBridge == null)
                    return;

                try
                {
                    RemoveEventHandler(_addInManagerBridge);
                }
                finally
                {
                    _addInManagerBridge = null;
                }
        }

        private void RemoveEventHandler(AddInManagerBridge addInManagerBridge)
        {
            //addInManagerBridge.TestSuiteRequest -= AddInBridgeTestSuiteRequest;
        }

        public void Dispose()
        {
                Dispose(true);
                GC.SuppressFinalize(this);
        }

        ~AddInManager()
        {
            Dispose(false);
        }

        #endregion
    }
}
