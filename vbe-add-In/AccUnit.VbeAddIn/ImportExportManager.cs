using System;
using System.Collections.Generic;
using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.AccUnit.VbeAddIn.Resources;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Commandbar;
using Microsoft.Office.Core;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class ImportExportManager : IDisposable, ICommandBarsAdapterClient
    {

        public event EventHandler TestClassesImported;

        public TestClassManager TestClassManager { get; set; }

        void ShowImportDialog()
        {
            try
            {
                using (var importForm = new TestClassSelectionForm(TestClassSelectionForm.SelectionMode.Import, GetTestTestModulesFromImportDirectory()))
                {
                    ShowTestClassSelectionDialog(importForm);
                }
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
            }
        }

        void ShowExportDialog()
        {
            try
            {
                var list = TestClassManager.GetTestModulesFromVBProject();
                using (var exportForm = new TestClassSelectionForm(TestClassSelectionForm.SelectionMode.Export, list))
                {
                    ShowTestClassSelectionDialog(exportForm);
                }
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
            }
        }

        void ShowTestClassSelectionDialog(TestClassSelectionForm form)
        {
            form.RefreshTestList += TestClassSelectionFormRefreshTestList;
            form.TestsSelected += TestClassSelectionFormTestsSelected;
            form.ShowDialog();
            form.RefreshTestList -= TestClassSelectionFormRefreshTestList;
            form.TestsSelected -= TestClassSelectionFormTestsSelected;
            form.Dispose();
        }

        void TestClassSelectionFormTestsSelected(TestClassSelectionForm sender, TestComponentsEventArgs e, ref bool close)
        {
            try
            {
                if (sender.CurrentMode == TestClassSelectionForm.SelectionMode.Export)
                    TestClassManager.RemoveTestComponents(e.Components);
                else
                {
                    TestClassManager.ImportTestComponents(e.Components);
                    if (TestClassesImported != null)
                        TestClassesImported(this, null);
                }
                close = true;
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
            }
        }

        void TestClassSelectionFormRefreshTestList(TestClassSelectionForm sender, TestComponentsEventArgs e)
        {
            e.Components = (sender.CurrentMode == TestClassSelectionForm.SelectionMode.Import ? GetTestTestModulesFromImportDirectory() : TestClassManager.GetTestModulesFromVBProject());
        }

        IEnumerable<CodeModuleInfo> GetTestTestModulesFromImportDirectory()
        {
            return TestClassManager.GetTestModulesFromDirectory();
        }

        #region ICommandBarsAdapterClient support

        public void SubscribeToCommandBarAdapter(VbeCommandBarAdapter commandBarAdapter)
        {
            using (new BlockLogger())
            {
                var accUnitCommandBarAdapter = commandBarAdapter as AccUnitCommandBarAdapter;
                if (accUnitCommandBarAdapter == null) return;

                var menu = accUnitCommandBarAdapter.AccUnitSubMenu;
                var buttonData = new CommandbarButtonData
                                 {
                                     Caption = VbeCommandbars.ToolsImportTestsCommandButtonCaption,
                                     Description = string.Empty,
                                     FaceId = 524,
                                     BeginGroup = true
                                 };
                commandBarAdapter.AddCommandBarButton(menu, buttonData, AccUnitMenuItemsImportTests);

                buttonData = new CommandbarButtonData
                             {
                                 Caption = VbeCommandbars.ToolsExportTestsCommandButtonCaption,
                                 Description = string.Empty,
                                 FaceId = 525,
                                 BeginGroup = false
                             };
                commandBarAdapter.AddCommandBarButton(menu, buttonData, AccUnitMenuItemsExportTests);
            }
        }

        void AccUnitMenuItemsImportTests(CommandBarButton ctrl, ref bool cancelDefault)
        {
            ShowImportDialog();
        }

        void AccUnitMenuItemsExportTests(CommandBarButton ctrl, ref bool cancelDefault)
        {
            ShowExportDialog();
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
        //    //
        //}

        private void DisposeManagedResources()
        {
            TestClassManager = null;
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~ImportExportManager()
        {
            Dispose(false);
        }

        #endregion
    }
}
