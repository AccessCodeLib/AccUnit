using System;
using System.Diagnostics;
using System.Reflection;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Integration;
using AccessCodeLib.AccUnit.VbeAddIn.Properties;
using Microsoft.Vbe.Interop;
using Timer = System.Windows.Forms.Timer;
using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.Generic;
using System.Linq;
using System.Collections;
using AccessCodeLib.AccUnit.VbeAddIn.TestExplorer;
using AccessCodeLib.AccUnit.Configuration;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    internal class AddInManager : IDisposable
    {
        private AddIn _addIn;
        private Timer _startupTimer;
        private OfficeApplicationHelper _officeApplicationHelper;
        
        private readonly VbeIntegrationManager _vbeIntegrationManager = new VbeIntegrationManager();
        private readonly TestSuiteManager _testSuiteManager = new TestSuiteManager();

        private readonly AccUnitCommandBarAdapter _commandBarsAdapter;
        private readonly TestStarter _testStarter = new TestStarter();

        private readonly TestExplorerManager _testExplorerManager;

        /*
        private readonly TagListManager _tagListManager = new TagListManager();
        private readonly TestListAndResultManager _testListAndResultManager = new TestListAndResultManager();
        
        private readonly DialogManager _dialogManager = new DialogManager();
        private readonly TestTemplateGenerator _testTemplateGenerator = new TestTemplateGenerator();
        private AccSpecCommandBarAdapterClient _accSpecCommandBarAdapterClient;
        private AccSpecManager _accSpecManager;

        
        */

        //private VbaProgrammingTools _vbaProgrammingTools;

        public AddInManager(AddIn addIn)
        {
            using (new BlockLogger())
            {
                _addIn = addIn;

                /*
                _tagListManager.AddIn = addIn;
                _testListAndResultManager.AddIn = addIn;
               

                if (Settings.Default.VbaProgrammingToolsEnabled)
                {
                    _vbaProgrammingTools = new VbaProgrammingTools();
                }
                */

                InitOfficeApplicationHelper();
                //InitAccSpec();

                var testExplorer = new TestExplorerView();
                var vbeControl = new VbeUserControl<TestExplorerView>(AddIn, "AccUnit Test Explorer", TestExplorerInfo.PositionGuid, testExplorer, false);
                _testExplorerManager = new TestExplorerManager(vbeControl);

                _commandBarsAdapter = new AccUnitCommandBarAdapter(VBE);
            }
        }

        /*
        private void InitAccSpec()
        {
            VbeManager vbeManager = null;
            var currentVbProject = _officeApplicationHelper.CurrentVBProject;
            if (currentVbProject != null)
                vbeManager = new VbeManager(currentVbProject);

            _accSpecManager = new AccSpecManager(vbeManager);
            _accSpecCommandBarAdapterClient = new AccSpecCommandBarAdapterClient(_accSpecManager);
            _vbeIntegrationManager.ScanningForTestModules += OnScanningForTestModules;
            _testStarter.ScanningForTestModules += OnScanningForTestModules;
        }
        */

        public void Init()
        {
            using (new BlockLogger())
            {
                if (!ApplicationSupportsAddIn)
                    return;

                try
                {
                    InitTestSuiteManager();
                    InitVbeWindows();
                    InitVbeIntegrationManager();

                    //_testListAndResultManager.TagListManager = _tagListManager;
                    //_testStarter.te = _testListAndResultManager;
                    _testExplorerManager.RunTests += OnRunTests;    
                    _testStarter.ShowUIMessage += OnShowUIMessage;

                    InitCommandBarsAdapter();
                }
                catch (Exception ex)
                {
                    UITools.ShowException(ex);
                }
            }
        }

        private void OnRunTests(object sender, RunTestsEventArgs e)
        {
            _testStarter.RunTests(e.TestClassList);
        }

        static void OnShowUIMessage(object sender, MessageEventArgs e)
        {
            UITools.ShowMessage(e.Message, e.Buttons, e.Icon, e.DefaultButton);
            e.MessageDisplayed = true;
        }

        private bool ApplicationSupportsAddIn => !(_officeApplicationHelper is VbeOnlyApplicatonHelper);

        private void InitCommandBarsAdapter()
        {
            using (new BlockLogger())
            {
                _commandBarsAdapter.Init();
                _commandBarsAdapter.AddClient(_testStarter);
                _commandBarsAdapter.AddClient(_vbeIntegrationManager);
                _commandBarsAdapter.AddClient(_testExplorerManager);
                
                /*
                _commandBarsAdapter.AddClient(_testListAndResultManager);
                _commandBarsAdapter.AddClient(_tagListManager);
                _commandBarsAdapter.AddClient(_testTemplateGenerator);
                
                _commandBarsAdapter.AddClient(_dialogManager);
                */
                /*
                if (UserSettings.Current.IsAccSpecEnabled)
                {
                    _commandBarsAdapter.AddClient(_accSpecCommandBarAdapterClient);
                }

                if (_vbaProgrammingTools != null)
                {
                    AddVbaProgrammingToolsToCommandBar();
                }
                */
            }
        }

        /*
        private void AddVbaProgrammingToolsToCommandBar()
        {
            using (new BlockLogger())
            {
                _commandBarsAdapter.AddClient(_vbaProgrammingTools);
            }
        }
        */

        private void InitTestSuiteManager()
        {
            using (new BlockLogger())
            {
                _testSuiteManager.TestResultReporterRequest += TestSuiteManager_TestResultReporterRequest;

                _testStarter.TestSuiteManager = _testSuiteManager;
                
                //_testListAndResultManager.TestSuiteManager = _testSuiteManager;
            }
        }

        private void TestSuiteManager_TestResultReporterRequest(ref IList<ITestResultReporter> reporters)
        {            
            var loggerControl = new LoggerControl();
            loggerControl.LogTextBox.AppendText("...");
            var vbeControl = new VbeUserControl<LoggerControl>(AddIn, "AccUnit Test Result Logger", LoggerControlInfo.PositionGuid, loggerControl);
            reporters.Add(new LoggerControlReporter(vbeControl));

            if (!reporters.Contains(_testExplorerManager))
            {
                reporters.Add(_testExplorerManager);
            }
        }

        private void InitVbeIntegrationManager()
        {
            using (new BlockLogger())
            {
                //_tagListManager.VbeIntegrationManager = _vbeIntegrationManager;
                //_testListAndResultManager.VbeIntegrationManager = _vbeIntegrationManager;
                _testExplorerManager.VbeIntegrationManager = _vbeIntegrationManager;
                _testStarter.VbeIntegrationManager = _vbeIntegrationManager;
                
                //_testTemplateGenerator.VbeIntegrationManager = _vbeIntegrationManager;

                _vbeIntegrationManager.VBProjectChanged += VbeIntegrationManagerOnVBProjectChanged;
            }
        }

        void VbeIntegrationManagerOnVBProjectChanged(object sender, VbProjectEventArgs e)
        {
            /*
            var accessTestSuite = _testSuiteManager.TestSuite as AccessTestSuite;
            if (accessTestSuite != null)
            {
                accessTestSuite.HostApplication = _officeApplicationHelper.Application;
            }
            else
            {
                var vbaTestSuite = (VBATestSuite)_testSuiteManager.TestSuite;
                vbaTestSuite.HostApplication = _officeApplicationHelper.Application;
                vbaTestSuite.ActiveVBProject = e.VBProject;
            }
            */
            TestClassManager.ApplicationHelper = _officeApplicationHelper;
            //_accSpecManager.VbeManager = new VbeManager(_officeApplicationHelper.CurrentVBProject);
            
        }

        void OnScanningForTestModules(object sender, EventArgs e)
        {
            /*
            if (UserSettings.Current.IsAccSpecEnabled)
            {
                _accSpecManager.TransformFeatures();
            }
            */
        }

        private TestClassManager TestClassManager => _vbeIntegrationManager.TestClassManager;

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
            using (new BlockLogger())
            {
                _addInManagerBridge = new AddInManagerBridge();
                _addInManagerBridge.TestSuiteRequest += AddInBridgeTestSuiteRequest;
                _addInManagerBridge.HostApplicationInitialized += AddInBridgeHostApplicationInitialized;
                _addInManagerBridge.ConstraintBuilderRequest += AddInBridgeConstraintBuilderRequest;    
                _addInManagerBridge.AssertRequest += AddInBridgeAssertRequest;  
            }
        }

        private void AddInBridgeAssertRequest(out Interop.IAssert assert)
        {
            assert = _testSuiteManager.Assert;
        }

        private void AddInBridgeConstraintBuilderRequest(out Interop.IConstraintBuilder constraintBuilder)
        {
            constraintBuilder = _testSuiteManager.ConstraintBuilder;    
        }

        void AddInBridgeTestSuiteRequest(out IVBATestSuite suite)
        {
            suite = _testSuiteManager.TestSuite;
        }

        void AddInBridgeHostApplicationInitialized(object application)
        {
            InitOfficeApplicationHelper(application);
        }

        private void InitOfficeApplicationHelper(object hostApplication = null)
        {
            using (new BlockLogger())
            {
                _officeApplicationHelper = HostApplicationTools.GetOfficeApplicationHelper(VBE, ref hostApplication);
                _vbeIntegrationManager.OfficeApplicationHelper = _officeApplicationHelper;
                _testSuiteManager.OfficeApplicationHelper = _officeApplicationHelper;

                /*
                if (_vbaProgrammingTools != null)
                {
                    InitVbaProgrammingTools(_officeApplicationHelper);
                }
                */
            }
        }

        #endregion

        private void InitVbaProgrammingTools(OfficeApplicationHelper officeApplicationHelper)
        {
            /*
            using (new BlockLogger())
            {
                _vbaProgrammingTools.OfficeApplicationHelper = officeApplicationHelper;
            }
            */
        }

        #region ad VbeWindow



        private void InitVbeWindows()
        {
            bool testListVisible;
            using (new BlockLogger("Getting testListVisible"))
            {
                // PERF: This takes long, consider retrieving the values later
                using (new BlockLogger("PERF: Reading some settings"))
                {
                    bool restoreVbeWindowsStateOnLoad;
                    using (new BlockLogger("PERF: Reading 1st setting from Settings.Default"))
                    {
                        restoreVbeWindowsStateOnLoad = Settings.Default.RestoreVbeWindowsStateOnLoad;
                    }
                    var listVisible = Settings.Default.TestListVisible;
                    testListVisible = restoreVbeWindowsStateOnLoad && listVisible;
                }
            }
            /*
            using (new BlockLogger("Settings.Default.TestListVisible = " + testListVisible))
            {
                if (!testListVisible) return;
                _testListAndResultManager.ShowTestListWindow(true, false);
                InitStartUpTimer(1000, true);
            }
            */
        }

        private void InitStartUpTimer(int interval, bool start)
        {
            if (_startupTimer == null)
            {
                _startupTimer = new Timer();
                _startupTimer.Tick += StartupTimerTick;
            }
            _startupTimer.Interval = interval;
            if (start) 
                _startupTimer.Start();
        }

        private void DisposeStartUpTimer()
        {
            if (_startupTimer == null)
                return;

            using (new BlockLogger())
            {
                try
                {
                    _startupTimer.Stop();
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }

                try
                {
                    _startupTimer.Tick -= StartupTimerTick;
                }
                catch(Exception ex)
                {
                    Logger.Log(ex);
                }

                try
                {
                    _startupTimer.Dispose();
                    _startupTimer = null;
                    Logger.Log("_startupTimer disposed");
                }
                catch(Exception ex)
                {
                    Logger.Log(ex);
                }
            }
        }

        void StartupTimerTick(object sender, EventArgs e)
        {
            _startupTimer.Stop();
            //_testListAndResultManager.AddTestClassListToTestListAndResultWindow();
            DisposeStartUpTimer();
        }

        #endregion

        public static string FriendlyName => $"AccUnit {FileVersion}";

        public static string FileVersion
        {
            get
            {
                var version = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                return version.FileVersion;
            }
        }

        private VBE VBE => AddIn.VBE;

        private Microsoft.Vbe.Interop.AddIn AddIn => _addIn;



        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            using (new BlockLogger())
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
                       Logger.Log(ex);
                    }
                }

                try
                {
                    DisposeUnManagedResources();
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                _disposed = true;
            }
        }

        private void DisposeUnManagedResources()
        {
            _addIn = null;
        }

        private void DisposeManagedResources()
        {
            using (new BlockLogger())
            {
                try
                {
                    DisposeStartUpTimer();
                    //_testListAndResultManager.Dispose();
                    DisposeVbaProgrammingTools();
                    /*
                    _testStarter.Dispose();

                    _testTemplateGenerator.Dispose();

                    _vbeIntegrationManager.Dispose();

                    try
                    {
                        _commandBarsAdapter.Dispose(); 
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(ex);
                    }
                    */

                    _testSuiteManager.Dispose();
                    
                    DisposeAddInManagerBridge();   

                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }

                try
                {
                    _officeApplicationHelper.Dispose();
                    _officeApplicationHelper = null;
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
            }
        }

        private void DisposeVbaProgrammingTools()
        {
            //_vbaProgrammingTools?.Dispose();
            //_vbaProgrammingTools = null;
        }

        private void DisposeAddInManagerBridge()
        {
            using (new BlockLogger())
            {
                if (_addInManagerBridge == null)
                    return;

                try
                {
                    RemoveEventHandler(_addInManagerBridge);
                    Logger.Log("_addInManagerBridge disposed");
                }
                finally
                {
                    _addInManagerBridge = null;
                }
            }
        }

        private void RemoveEventHandler(AddInManagerBridge addInManagerBridge)
        {
            addInManagerBridge.TestSuiteRequest -= AddInBridgeTestSuiteRequest;
            addInManagerBridge.HostApplicationInitialized -= AddInBridgeHostApplicationInitialized;
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~AddInManager()
        {
            Logger.Log("~AddInManager");
            Dispose(false);
            Logger.Log("~AddInManager completed");
        }

        #endregion
    }
}
