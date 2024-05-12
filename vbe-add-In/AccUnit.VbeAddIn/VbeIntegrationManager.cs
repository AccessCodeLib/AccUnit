using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.AccUnit.Tools;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Commandbar;
using AccessCodeLib.Common.VBIDETools.Templates;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class VbeIntegrationManager : IDisposable, ICommandBarsAdapterClient
    {
        private readonly VbeAdapter _vbeAdapter = new VbeAdapter();
        private readonly ImportExportManager _importExportManager = new ImportExportManager();
        private TestClassManager _testClassManager;

        public VbeIntegrationManager()
        {
            _importExportManager.TestClassesImported += ImportExportManagerTestClassesImported;
        }

        public event EventHandler<VbProjectEventArgs> VBProjectChanged;
        public event EventHandler ScanningForTestModules;

        public OfficeApplicationHelper OfficeApplicationHelper
        {
            get { return _vbeAdapter.OfficeApplicationHelper; }
            set
            {
                _vbeAdapter.OfficeApplicationHelper = value;
                _importExportManager.TestClassManager = TestClassManager;
            }
        }

        public VbeAdapter VbeAdapter
        {
            get { return _vbeAdapter; }
        }

        void ImportExportManagerTestClassesImported(object sender, EventArgs e)
        {
            if (!(OfficeApplicationHelper is AccessApplicationHelper)) return;
            ((AccessApplicationHelper)OfficeApplicationHelper).RunCommand(AccessApplicationHelper.AcCommand.AcCmdCompileAndSaveAllModules);
        }

        private VBProject ActiveVBProject
        {
            get { return _vbeAdapter.ActiveVBProject; }
        }

        public TestClassManager TestClassManager
        {
            get
            {
                if (_testClassManager == null)
                    InitTestClassManager();

                return _testClassManager;
            }
        }

        private void InitTestClassManager()
        {
            _testClassManager = new TestClassManager(OfficeApplicationHelper);
            _importExportManager.TestClassManager = _testClassManager;
            _testClassManager.RepairActiveVBProjectCOMException += TestClassManagerOnRepairActiveVbProjectComException;
            _testClassManager.ScanningForTestModules += TestClassManagerOnScanningForTestModules;
        }

        void TestClassManagerOnScanningForTestModules(object sender, EventArgs e)
        {
            RaiseScanningForTestModules();
        }

        public object HostApplication
        {
            get { return OfficeApplicationHelper.Application; }
        }

        private void SetTestEnvironment()
        {
            var configurator = new Configurator(ActiveVBProject);
            configurator.InsertAccUnitLoaderFactoryModule(AccUnitTypeLibIsReferenced, true);
        }

        private bool AccUnitTypeLibIsReferenced
        {
            get
            {
                foreach (Reference reference in ActiveVBProject.References)
                {
                    if (reference.IsBroken) continue;
                    if (reference.Name == "AccUnit")
                        return true;
                }
                return false;
            }
        }

        private void RemoveTestEnvironment()
        {
            if (MessageBox.Show(Resources.UserControls.RemoveEnvironmentMessageBoxText,
                                Resources.UserControls.RemoveEnvironmentMessageBoxCaption,
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Exclamation) != DialogResult.Yes) return;

            TestClassManager.RemoveTestComponents(true, true);
        }

        private void CreateTestMethodInActiveCodePane()
        {
            var insertTestMethodDataContext = new InsertTestMethodViewModel();
            var dialog = new InsertTestMethodDialog(insertTestMethodDataContext);

            insertTestMethodDataContext.InsertTestMethod += (sender, e) =>
            {
                InsertTestMethodDialogCommitMethodName(sender, e);
                dialog.Close();
            };
            insertTestMethodDataContext.Canceled += (sender, e) => dialog.Close();

            SetDialogPosition(dialog);
            dialog.ShowDialog();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "default interface, but not used")]
        private void InsertTestMethodDialogCommitMethodName(InsertTestMethodViewModel sender, TestNamePartsEventArgs e)
        {
            var methodUnderTest = e.Items.FirstOrDefault(i => i.Name == InsertTestMethodViewModel.TestNamePart_MethodName)?.Value;
            var stateUnderTest = e.Items.FirstOrDefault(i => i.Name == InsertTestMethodViewModel.TestNamePart_State)?.Value;
            var expectedBehaviour = e.Items.FirstOrDefault(i => i.Name == InsertTestMethodViewModel.TestNamePart_Expected)?.Value;
            CreateTestMethodInActiveCodePane(methodUnderTest, stateUnderTest, expectedBehaviour);
        }

        private void CreateTestMethodInActiveCodePane(string methodUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            using (new BlockLogger())
            {
                if (string.IsNullOrEmpty(methodUnderTest))
                    methodUnderTest = "MethodUnderTest";
                var memberInfo = new CodeModuleMember(methodUnderTest, vbext_ProcKind.vbext_pk_Proc, true);
                var methodText = TestCodeGenerator.GenerateProcedureCode(memberInfo, stateUnderTest, expectedBehaviour);

                var activeCodePane = _vbeAdapter.ActiveCodePane;
                var startLine = VbeCodePaneTools.GetStartLineFromCodePaneSelection(activeCodePane);

                //Todo: check if startLine is inside a method   

                VbeCodePaneTools.InsertText(activeCodePane, methodText, startLine);
            }
        }

        public TestClassInfo GetTestClassInfoFromSelectedComponent()
        {
            var selectedComponent = _vbeAdapter.VBE.SelectedVBComponent;
            if (selectedComponent == null)
                return null; // no active test

            if (selectedComponent.Type != vbext_ComponentType.vbext_ct_ClassModule)
                return null;

            if (!TestClassReader.IsTestClassCodeModul(selectedComponent.CodeModule))
                return null;

            var className = selectedComponent.Name;
            var memberInfo = GetTestClassMemberInfoFromCodePane(selectedComponent.CodeModule.CodePane);

            TestClassInfo testclass;
            if (memberInfo != null)
            {
                var members = new TestClassMemberList { memberInfo };
                testclass = new TestClassInfo(className, members);
            }
            else
            {
                testclass = new TestClassInfo(className);
            }
            return testclass;
        }

        private static TestClassMemberInfo GetTestClassMemberInfoFromCodePane(_CodePane codepane)
        {
            var procName = VbeCodePaneTools.GetCodeModuleMemberNameFromCodePane(codepane, out vbext_ProcKind procKind);

            if (procKind == vbext_ProcKind.vbext_pk_Proc && procName != null)
            {
                return new TestClassMemberInfo(procName);
            }
            return null;
        }

        private void InsertTestMethodDialogCommitMethodName(object sender, CommitInsertTestMethodEventArgs e)
        {
            CreateTestMethodInActiveCodePane(e.MethodUnderTest, e.StateUnderTest, e.ExpectedBehaviour);
        }

        private void TestClassManagerOnRepairActiveVbProjectComException(object sender, EventArgs e)
        {
            try
            {
                RepairTestSuiteCOMException();
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        private void RaiseScanningForTestModules()
        {
            ScanningForTestModules?.Invoke(this, EventArgs.Empty);
        }

        private void RepairTestSuiteCOMException()
        {
            VBProjectChanged?.Invoke(this, new VbProjectEventArgs(OfficeApplicationHelper.CurrentVBProject));

            TestClassManager.ApplicationHelper = OfficeApplicationHelper;
        }

        public void ShowMissingApplicationInfo()
        {
            if (HostApplication == null)
            {
                UITools.ShowMessage(Resources.MessageStrings.MissingHostApplicationReference);
            }
        }

        #region ICommandBarsAdapterClient support

        public void SubscribeToCommandBarAdapter(VbeCommandBarAdapter commandBarAdapter)
        {
            using (new BlockLogger())
            {
                _importExportManager.SubscribeToCommandBarAdapter(commandBarAdapter);
                CreateShortcutMenuItems(commandBarAdapter);
            }
        }

        private void CreateShortcutMenuItems(VbeCommandBarAdapter commandBarAdapter)
        {
            if (commandBarAdapter is AccUnitCommandBarAdapter accUnitCommandBarAdapter)
            {
                var menu = accUnitCommandBarAdapter.AccUnitSubMenu;
                CreateAccUnitToolsSubMenuItems(commandBarAdapter, menu);
            }

            var commandBar = commandBarAdapter.CommandBarCodeWindow;
            const int objectBrowserControlID = 473;
            CreateAccUnitShortcutMenuItems(commandBarAdapter, commandBar, objectBrowserControlID);

            commandBar = commandBarAdapter.CommandBarProjectWindow;
            const int printControlID = 4;
            CreateAccUnitShortcutMenuItems(commandBarAdapter, commandBar, printControlID);

        }

        private void CreateAccUnitToolsSubMenuItems(VbeCommandBarAdapter commandBarAdapter, CommandBarPopup menu)
        {
            CreateAccUnitToolsSetTestEnvironmentSubMenuItem(commandBarAdapter, menu);
            CreateAccUnitToolsRemoveTestEnvironmentSubMenuItem(commandBarAdapter, menu);
        }

        private void CreateAccUnitToolsSetTestEnvironmentSubMenuItem(VbeCommandBarAdapter commandBarAdapter, CommandBarPopup menu)
        {
            var buttonData = new CommandbarButtonData
            {
                Caption = Resources.VbeCommandbars.ToolsSetTestEnvironmentCommandButtonCaption,
                Description = string.Empty,
                FaceId = 589,
                BeginGroup = true
            };
            commandBarAdapter.AddCommandBarButton(menu, buttonData, AccUnitMenuItemsSetTestEnvironment);
        }

        private void CreateAccUnitToolsRemoveTestEnvironmentSubMenuItem(VbeCommandBarAdapter commandBarAdapter, CommandBarPopup menu)
        {
            var buttonData = new CommandbarButtonData
            {
                Caption = Resources.VbeCommandbars.ToolsRemoveTestEnvironmentCommandButtonCaption,
                Description = string.Empty,
                FaceId = 478,
                BeginGroup = false
            };
            commandBarAdapter.AddCommandBarButton(menu, buttonData, AccUnitMenuItemsRemoveTestEnvironment);
        }

        private void CreateAccUnitShortcutMenuItems(VbeCommandBarAdapter accUnitMenuItems, CommandBar commandBar, int controlID)
        {
            var objectBrowserControlIndex = VbeCommandBarAdapter.GetButtonIndex(commandBar, controlID);
            var buttonData = new CommandbarButtonData
            {
                Caption = Resources.VbeCommandbars.InsertTestMethodCommandbarButtonCaption,
                Description = string.Empty,
                FaceId = 559,
                BeginGroup = true,
                Index = objectBrowserControlIndex
            };
            accUnitMenuItems.AddCommandBarButton(commandBar, buttonData, AccUnitMenuItemsInsertTestMethod);
        }

        void AccUnitMenuItemsSetTestEnvironment(CommandBarButton ctrl, ref bool cancelDefault)
        {
            SetTestEnvironment();
        }

        void AccUnitMenuItemsRemoveTestEnvironment(CommandBarButton ctrl, ref bool cancelDefault)
        {
            RemoveTestEnvironment();
        }

        void AccUnitMenuItemsInsertTestMethod(CommandBarButton ctrl, ref bool cancelDefault)
        {
            if (SelectedVbComponentIsTestClass)
                CreateTestMethodInActiveCodePane();
            else
                CreateTestMethodFromSelectedVbComponent();
        }

        private bool SelectedVbComponentIsTestClass
        {
            get
            {
                return TestClassReader.IsTestClassCodeModul(_vbeAdapter.VBE.SelectedVBComponent.CodeModule);
            }
        }

        public object VBETools { get; private set; }

        private void CreateTestMethodFromSelectedVbComponent()
        {
            var generateTestMethodsFromCodeModuleDataContext = new GenerateTestMethodsFromCodeModuleViewModel(GetCodeModuleInfoWithMarkerFromSelectedVbComponent());
            var dialog = new GenerateTestMethodsFromCodeModuleDialog(generateTestMethodsFromCodeModuleDataContext);

            generateTestMethodsFromCodeModuleDataContext.InsertTestMethods += (sender, e) =>
            {
                var newCodeModule = InsertTestMethodsDialogCommitMethodName(sender, e);
                dialog.Close();
                if (newCodeModule != null)
                {
                    newCodeModule.CodePane.Show();
                    newCodeModule.CodePane.Window.SetFocus();
                }
            };
            generateTestMethodsFromCodeModuleDataContext.Canceled += (sender, e) => dialog.Close();

            SetDialogPosition(dialog);
            dialog.ShowDialog();
        }

        private void SetDialogPosition(System.Windows.Window dialog)
        {
            var scaleFactor = UITools.GetScalingFactor();
            var width = dialog.MinWidth;
            var height = dialog.MaxHeight;

            dialog.Top = (_vbeAdapter.VBE.MainWindow.Top + _vbeAdapter.VBE.MainWindow.Height / 2) / scaleFactor - height / 2;
            dialog.Left = (_vbeAdapter.VBE.MainWindow.Left + _vbeAdapter.VBE.MainWindow.Width / 2) / scaleFactor - width / 2;
        }

        private CodeModule InsertTestMethodsDialogCommitMethodName(object sender, CommitInsertTestMethodsEventArgs e)
        {
            var testClassGenerator = new TestClassGenerator(ActiveVBProject);
            try
            {
                using (new BlockLogger($"{e.TestClass}.{e.MethodsUnderTest}_{e.StateUnderTest}_{e.ExpectedBehaviour}"))
                {
                    return testClassGenerator.InsertTestMethods(e.CodeModuleToTest, e.TestClass, e.MethodsUnderTest, e.StateUnderTest, e.ExpectedBehaviour);
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                e.Cancel = true;
                return null;
            }
        }

        private CodeModuleInfo GetCodeModuleInfoWithMarkerFromSelectedVbComponent()
        {
            var vbc = _vbeAdapter.VBE.SelectedVBComponent;
            var reader = new CodeModuleReader(vbc.CodeModule);
            var codemoduleInfo = reader.CodeModuleInfo;

            var codePane = _vbeAdapter.ActiveCodePane;
            var activeMemberName = (codePane == null || codePane.CodeModule != vbc.CodeModule) ? null : VbeCodePaneTools.GetCodeModuleMemberNameFromCodePane(codePane);

            var markAll = (activeMemberName == null);
            var publicMembers = codemoduleInfo.Members.FindAll(m => m.IsPublic);

            var newMembers = new CodeModuleMemberList();
            newMembers.AddRange(publicMembers.Select(newMember => new CodeModuleMemberWithMarker(newMember.Name, newMember.ProcKind, newMember.IsPublic, newMember.DeclarationString, markAll)).Cast<CodeModuleMember>());
            if (!markAll)
            {
                var markedMember = (CodeModuleMemberWithMarker)newMembers.Find(m => m.Name == activeMemberName);
                if (markedMember != null) markedMember.Marked = true;
                var info = activeMemberName;
                if (markedMember != null) info += string.Format(": {0}", markedMember.Marked);
                Logger.Log(info);
            }

            codemoduleInfo.Members = newMembers;
            return codemoduleInfo;
        }

        public void InsertTestTemplate(string templatekey)
        {
            var templates = UserSettings.Current.TestTemplates;
            var template = templates[templatekey];
            var name = FindFreeClassName(template.Name);
            do
            {
                if (DialogResult.OK != UITools.InputBox(Resources.UserControls.InsertTestTemplateInputboxTitle,
                                                    Resources.UserControls.InsertTestTemplateInputboxPromptText,
                                                    ref name))
                    return;

            } while (!InsertTestTemplate(template, name));
        }

        private string FindFreeClassName(string defaultName)
        {
            var codeModuleContainer = new CodeModuleContainer(ActiveVBProject);
            var newClassName = defaultName;
            var i = 0;
            while (codeModuleContainer.Exists(newClassName))
            {
                newClassName = string.Format("{0}{1}", defaultName, ++i);
            }
            return newClassName;
        }

        private bool InsertTestTemplate(CodeTemplate template, string className)
        {
            try
            {
                TestClassManager.InsertTestTemplate(template, className);
                return true;
            }
            catch (ArgumentException ex)
            {
                UITools.ShowMessage(ex.Message);
                return false;
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
                return false;
            }
        }

        #endregion

        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                DisposeManagedResources();
            }

            //DisposeUnmanagedResources();
            _disposed = true;
        }

        //private void DisposeUnmanagedResources()
        //{
        //}

        private void DisposeManagedResources()
        {
            if (_testClassManager != null)
            {
                _testClassManager.Dispose();
                _testClassManager = null;
            }

            _importExportManager.Dispose();
            _vbeAdapter.Dispose();
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~VbeIntegrationManager()
        {
            Dispose(false);
        }

        #endregion
    }
}
