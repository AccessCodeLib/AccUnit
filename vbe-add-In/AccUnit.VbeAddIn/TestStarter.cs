using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Commandbar;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class TestStarter : IDisposable, ICommandBarsAdapterClient
    {
        bool _referencesChecked;
        private bool _breakOnAllErrorsForNextRun;
        private CommandBarEvents _accUnitSubMenuEvents;
        private int _accUnitSubMenuRunCurrentTestIndex;
        private CommandBarButton _codeWindowRunCurrentTestButton;
        private CommandBarButton _projectWindowRunCurrentTestButton;

        public VbeIntegrationManager VbeIntegrationManager { get; set; }

        void OnVbeMainWindowRButtonDown(object sender, EventArgs e)
        {
            try
            {
                var enabled = SelectedCodeModuleIsTestClass;
                _codeWindowRunCurrentTestButton.Enabled = enabled;
                _projectWindowRunCurrentTestButton.Enabled = enabled;
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        private TestClassManager TestClassManager { get { return VbeIntegrationManager.TestClassManager; } }
        //private object HostApplication { get { return VbeIntegrationManager.HostApplication; } }
        private VBE VBE { get { return VbeIntegrationManager.OfficeApplicationHelper.VBE; } }

        public TestSuiteManager TestSuiteManager { get; set; }
        private IVBATestSuite TestSuite { get { return TestSuiteManager.TestSuite; } }

        public event EventHandler<MessageEventArgs> ShowUIMessage;
        public event EventHandler ScanningForTestModules;

        private void RunCurrentTests(ResetMode resetmode = ResetMode.RemoveTests | ResetMode.ResetTestSuite)
        {
            try
            {
                TestClassInfo testclass = VbeIntegrationManager.GetTestClassInfoFromSelectedComponent();
                if (testclass != null)
                {
                    var list = new TestClassList { testclass };
                    RunTests(list, resetmode);
                }
            }
            catch (Exception xcp)
            {
                UITools.ShowException(xcp);
            }
        }

        public void RunAllTests(ResetMode resetmode = ResetMode.RemoveTests)
        {
            try
            {
                RunTests(null, resetmode);
            }
            catch (Exception ex)
            {
                if (ShowMessage(ex))
                    return;

                throw;
            }
        }

        public void RunTests(ICollection<TestClassInfo> testClassList, bool BreakOnAllErrors = false)
        {
            try
            {
                if (testClassList.Count > 0)
                {
                    _breakOnAllErrorsForNextRun = BreakOnAllErrors;
                    var missingTestClass = TestClassManager.FindFirstMissingTestClassInVBProject(testClassList);
                    if (missingTestClass != null)
                    {
                        UITools.ShowMessage(string.Format(Resources.MessageStrings.MissingTestClassInVBProject, missingTestClass.Name));
                        return;
                    }
                    RunTests(testClassList, ResetMode.RemoveTests);
                }
            }
            catch (Exception ex)
            {
                if (ShowMessage(ex))
                    return;

                throw;
            }
        }

        private bool ShowMessage(Exception ex)
        {
            return ShowMessage(ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private bool ShowMessage(string message, MessageBoxButtons buttons = MessageBoxButtons.OK,
                                 MessageBoxIcon icon = MessageBoxIcon.Information,
                                 MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
        {
            if (ShowUIMessage != null)
            {
                var e = new MessageEventArgs(message, buttons, icon, defaultButton);
                ShowUIMessage(this, e);
                if (e.MessageDisplayed)
                    return true;
            }
            return false;
        }

        private void RunTests(IEnumerable<TestClassInfo> list, ResetMode resetmode)
        {
            var testSuite = TestSuite.Reset(resetmode) as IVBATestSuite;
            CheckReferences();

            if (testSuite is AccessTestSuite accessSuite)
            {
                // TODO ScanningForTestModules: This triggers the ScanningForTestModules event, checkout if necessary
                if (!CheckAccessApplicationIsCompiledAndRefreshFactoryModule(accessSuite))
                    return;

                if (_breakOnAllErrorsForNextRun)
                {
                    accessSuite.ErrorTrapping = VbaErrorTrapping.BreakOnAllErrors;
                    _breakOnAllErrorsForNextRun = false;
                }
            }

            // TODO ScanningForTestModules: This triggers the ScanningForTestModules event, checkout if necessary
            AddTests(testSuite, list, resetmode);

            testSuite.Run();
            //Task.Run(() => testSuite.Run());
        }

        private bool CheckAccessApplicationIsCompiledAndRefreshFactoryModule(AccessTestSuite accessSuite)
        {
            if (accessSuite.CheckAccessApplicationIsCompiled()) return true;

            // safety: refresh factory procedures (if class renamed)
            var saveModules = false;
            try
            {
                saveModules = accessSuite.ActiveVBProject.Saved;
            }
            catch { }

            try // issue #37
            {
                RaiseScanningForTestModules();
                accessSuite.Reset(ResetMode.RefreshFactoryModule);
            }
            catch
            {
                //accessSuite.HostApplication = HostApplication;
                accessSuite.Reset(ResetMode.RefreshFactoryModule);
            }
            finally
            {
                if (saveModules)
                    TryExecuteSaveMenuItem();
            }

            try // issue #68 (try to compile)
            {
                if (VbeIntegrationManager.OfficeApplicationHelper is AccessApplicationHelper access)
                {
                    access.RunCommand(AccessApplicationHelper.AcCommand.AcCmdCompileAndSaveAllModules);
                    if (access.IsCompiled)
                        return true;
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }

            return DialogResult.Cancel != UITools.ShowMessage(Resources.MessageStrings.Application_not_saved_in_compiled_state,
                                                              MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2
                                              );
        }

        private void TryExecuteSaveMenuItem()
        {
            using (new BlockLogger())
            {
                try
                {
                    new VbeCommandBarAdapter(VBE).MenuBar.FindControl(MsoControlType.msoControlButton, 3, Type.Missing, Type.Missing, true).Execute();
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
            }
        }

        /*
        private void AddTestsAndTryToRepairComException(IVBATestSuite testSuite, IEnumerable<TestClassInfo> list, ResetMode resetmode)
        {
            try
            {
                AddTests(testSuite, list, resetmode);
            }
            catch (COMException comEx)  // issue #37
            {
                Logger.Log(comEx);
                RepairTestSuiteCOMException();
                AddTests(testSuite, list, resetmode);
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                if (HostApplication == null)
                {
                    throw new MissingHostException();
                }
                throw;
            }
        }

        private void RepairTestSuiteCOMException()
        {
            if (TestSuite is AccessTestSuite)
            {
                var accessSuite = (AccessTestSuite)TestSuite;
                accessSuite.HostApplication = HostApplication;
            }
            else
            {
                var vbaTestSuite = (VBATestSuite)TestSuite;
                vbaTestSuite.HostApplication = HostApplication;
                vbaTestSuite.ActiveVBProject = VbeIntegrationManager.OfficeApplicationHelper.CurrentVBProject;
            }
        }
        */

        private void AddTests(IVBATestSuite testSuite, IEnumerable<TestClassInfo> list, ResetMode resetmode)
        {

            if (list != null)
            {
                testSuite.AddTestClasses(list);
            }
            else if (resetmode == ResetMode.RemoveTests)
            {
                // TODO ScanningForTestModules: Call into TestClassManager.GetTestClassListFromVBProject() or alike - see comment block at the end of that method
                RaiseScanningForTestModules();
                testSuite.AddFromVBProject();
            }

            // like this (but maintain condition logic as done above, the following condition is not sufficient!)
            //if (list == null)
            //{
            //    var testClassManager = new TestClassManager();
            //    list = testClassManager.GetTestClassListFromVBProject(true);
            //}
            //testSuite.AddTestClasses(list);
        }

        private void RaiseScanningForTestModules()
        {
            ScanningForTestModules?.Invoke(this, EventArgs.Empty);
        }

        public void SubscribeToCommandBarAdapter(VbeCommandBarAdapter commandBarAdapter)
        {
            using (new BlockLogger())
            {
                AddRunButtonsToProjectExplorerContextMenu(commandBarAdapter);
                AddRunButtonsToCodeWindowContextMenu(commandBarAdapter);

                if (commandBarAdapter is AccUnitCommandBarAdapter accUnitCommandbarAdapter)
                {
                    AddRunButtonsToAccUnitSubMenu(accUnitCommandbarAdapter);
                    AddRunButtonsToAccUnitCommandBar(accUnitCommandbarAdapter);
                }

                VbeIntegrationManager.VbeAdapter.MainWindowRButtonDown += OnVbeMainWindowRButtonDown;

                RegisterHotKeys();
            }
        }

        private void AddRunButtonsToAccUnitSubMenu(AccUnitCommandBarAdapter commandBarAdapter)
        {
            var commandBar = commandBarAdapter.AccUnitSubMenu.CommandBar;
            CreateCommandBarItems(commandBarAdapter, commandBar, null, false);
            _accUnitSubMenuRunCurrentTestIndex = 1;
            _accUnitSubMenuEvents =
                commandBarAdapter.VBE.Events.CommandBarEvents[commandBarAdapter.AccUnitSubMenu];
            _accUnitSubMenuEvents.Click += OnAccUnitSubMenuEventsClick;
        }

        private void AddRunButtonsToProjectExplorerContextMenu(VbeCommandBarAdapter commandBarAdapter)
        {
            const int printControlID = 4;
            var commandBar = commandBarAdapter.CommandBarProjectWindow;
            var printControlIndex = VbeCommandBarAdapter.GetButtonIndex(commandBar, printControlID);
            _projectWindowRunCurrentTestButton = CreateCommandBarItems(commandBarAdapter, commandBar, printControlIndex, false);
        }

        private void AddRunButtonsToCodeWindowContextMenu(VbeCommandBarAdapter commandBarAdapter)
        {
            const int objectBrowserControlID = 473;
            var commandBar = commandBarAdapter.CommandBarCodeWindow;
            var objectBrowserControlIndex = VbeCommandBarAdapter.GetButtonIndex(commandBar, objectBrowserControlID);
            _codeWindowRunCurrentTestButton = CreateCommandBarItems(commandBarAdapter, commandBar, objectBrowserControlIndex, false);
        }

        private void AddRunButtonsToAccUnitCommandBar(AccUnitCommandBarAdapter commandBarAdapter)
        {
            var commandBar = commandBarAdapter.AccUnitCommandbar;
            CreateCommandBarItems(commandBarAdapter, commandBar, null, false);
            // @todo: check usePicture = true ... current: error in ApplyMaskedPicture
        }

        void OnAccUnitSubMenuEventsClick(object commandBarControl, ref bool handled, ref bool cancelDefault)
        {
            if (!(commandBarControl is CommandBarPopup mnu))
                return;

            mnu.Controls[_accUnitSubMenuRunCurrentTestIndex].Enabled = ActiveCodeModuleIsTestClass;
        }

        private bool ActiveCodeModuleIsTestClass
        {
            get
            {
                try
                {
                    var cm = VBE.ActiveCodePane.CodeModule;
                    return TestClassReader.IsTestClassCodeModul(cm);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                    return false;
                }
            }
        }

        private bool SelectedCodeModuleIsTestClass
        {
            get
            {
                try
                {
                    var component = VBE.SelectedVBComponent;
                    return component != null && TestClassReader.IsTestClassCodeModul(component.CodeModule);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                    return false;
                }
            }
        }

        private CommandBarButton CreateCommandBarItems(VbeCommandBarAdapter commandBarAdapter,
                                                       CommandBar commandBar,
                                                       int? positionBefore,
                                                       bool usePicture)
        {
            var runCurrentTestButton = AddRunCurrentTestsCommandBarButton(commandBarAdapter, commandBar, positionBefore, usePicture);
            AddRunAllTestsCommandBarButton(commandBarAdapter, commandBar, runCurrentTestButton.Index + 1, usePicture);

            return runCurrentTestButton;
        }

        private CommandBarButton AddRunCurrentTestsCommandBarButton(VbeCommandBarAdapter commandBarAdapter,
                                                                    CommandBar commandBar,
                                                                    int? positionBefore,
                                                                    bool usePicture)
        {
            var buttonData = new CommandbarButtonData
            {
                Caption = Resources.VbeCommandbars.RunCurrentTestCommandBarButtonCaption,
                Description = Resources.VbeCommandbars.RunCurrentTestCommandBarButtonDescription,
                FaceId = 2997,
                Index = positionBefore,
                BeginGroup = true
            };
            var button = commandBarAdapter.AddCommandBarButton(commandBar, buttonData, RunCurrentTestsEventHandler);
            button.ShortcutText = "Ctrl+Shift+T";
            /*
            if (usePicture)
            {
                ApplyMaskedPicture(button, Resources.Icons.runtest, Resources.Icons.runtest_mask);
            }
            */
            return button;
        }

        private CommandBarButton AddRunAllTestsCommandBarButton(VbeCommandBarAdapter commandBarAdapter,
                                                                CommandBar commandBar,
                                                                int? positionBefore,
                                                                bool usePicture)
        {
            var buttonData = new CommandbarButtonData
            {
                Caption = Resources.VbeCommandbars.RunAllTestsCommandBarButtonCaption,
                Description = Resources.VbeCommandbars.RunAllTestsCommandBarButtonDescription,
                FaceId = 620,
                Index = positionBefore
            };
            var button = commandBarAdapter.AddCommandBarButton(commandBar, buttonData, RunAllTestsEventHandler);
            button.ShortcutText = "Ctrl+Shift+A";
            /*
            if (usePicture)
            {
                ApplyMaskedPicture(button, Resources.Icons.runtests, Resources.Icons.runtests_mask);
            }
            */
            return button;
        }
        /*
        private void ApplyMaskedPicture(CommandBarButton button, Bitmap image, Bitmap mask)
        {
            try
            {
                using (new BlockLogger())
                {
                    button.Style = MsoButtonStyle.msoButtonAutomatic;
                    // PERF: The first conversions takes long
                    using (new BlockLogger("PERF: 1st call to ImageToPictureDisp takes long"))
                    {
                        button.Picture = AxHostConverter.ImageToPictureDisp(image);
                    }
                    button.Mask = AxHostConverter.ImageToPictureDisp(mask);
                }
            }
            catch (AccessViolationException e)
            {
                Logger.Log(e);
            }
        }*/

        private void RegisterHotKeys()
        {
            var hotkeys = VbeIntegrationManager.VbeAdapter.HotKeys;
            var hotKey = hotkeys.RegisterHotKey(HotKey.ModKeys.Control | HotKey.ModKeys.Shift, (uint)Keys.T);
            hotKey.Pressed += RunCurrentTestsHotKeyPressed;

            hotKey = hotkeys.RegisterHotKey(HotKey.ModKeys.Control | HotKey.ModKeys.Shift, (uint)Keys.A);
            hotKey.Pressed += RunAllTestsHotKeyPressed;
        }

        private void RunCurrentTestsHotKeyPressed(object sender, HotKeyEventArgs e)
        {
            RunCurrentTests();
        }

        private void RunAllTestsHotKeyPressed(object sender, HotKeyEventArgs e)
        {
            RunAllTests();
        }

        private void RunCurrentTestsEventHandler(CommandBarButton ctrl, ref bool cancelDefault)
        {
            RunCurrentTests();
        }

        private void RunAllTestsEventHandler(CommandBarButton ctrl, ref bool cancelDefault)
        {
            RunAllTests();
        }

        private void CheckReferences(bool constrainedCheck = false)
        {

            if (!constrainedCheck && _referencesChecked)
                return;
            /*
            using (var configurator = new Configurator())
            {
                configurator.AccUnitReference.EnsureReferenceExistsInVbProject(VBE.ActiveVBProject);
            }
            */
            _referencesChecked = true;
        }

        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                _codeWindowRunCurrentTestButton = null;
                _projectWindowRunCurrentTestButton = null;

                _accUnitSubMenuEvents = null;

                if (VbeIntegrationManager != null)
                {
                    try
                    {
                        VbeIntegrationManager.VbeAdapter.MainWindowRButtonDown -= OnVbeMainWindowRButtonDown;
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(ex);
                    }
                    finally
                    {
                        VbeIntegrationManager = null;
                    }
                }

                TestSuiteManager = null;
            }

            _disposed = true;
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~TestStarter()
        {
            Dispose(false);
        }

        #endregion
    }
}
