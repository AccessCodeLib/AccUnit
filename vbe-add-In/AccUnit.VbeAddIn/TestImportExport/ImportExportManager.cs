using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.AccUnit.VbeAddIn.Resources;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Commandbar;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
                var dataContext = new SelectControlViewModel(
                                                        UserControls.TestClassSelectionFormCaptionImport,
                                                        UserControls.SelectListSelectAllCheckboxCaption,
                                                        UserControls.TestClassSelectionFormCommitTestImport,
                                                        true, "Overwrite existing codemodule");

                foreach (var item in TestClassManager.GetTestModulesFromDirectory())
                {
                    dataContext.Items.Add(new CheckableCodeModuleInfo(item));
                }

                var form = new ImportExportWindow(dataContext);
                dataContext.OptionalCheckboxChecked = true;
                dataContext.RefreshList += (sender, e) =>
                {
                    e.Items.Clear();
                    foreach (var item in TestClassManager.GetTestModulesFromDirectory())
                    {
                        e.Items.Add(new CheckableCodeModuleInfo(item));
                    }
                };
                dataContext.ItemsSelected += (sender, e) =>
                {
                    var sb = new StringBuilder();
                    bool errRaised = false;

                    try
                    {
                        IEnumerable<CodeModuleInfo> codeModulesToImport = e.Items.Select(x => ((CheckableCodeModuleInfo)x).CodeModule);
                        if (codeModulesToImport.Count() == 0)
                        {
                            throw new Exception("No test modules selected for import.");
                        }
                        TestClassManager.ImportTestComponents(codeModulesToImport, dataContext.OptionalCheckboxChecked);
                    }
                    catch (Exception ex)
                    {
                        errRaised = true;
                        UITools.ShowException(ex);
                    }
                    TestClassesImported?.Invoke(this, null);

                    if (!errRaised)
                    {
                        UITools.ShowMessage(MessageStrings.TestImportedCommitMessage);
                        form.Close();
                    }
                };

                SetDialogPosition(form);
                form.ShowDialog();

                dataContext.RefreshList -= (sender, e) => { };
                dataContext.ItemsSelected -= (sender, e) => { };
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
                var dataContext = new SelectControlViewModel(
                                                        UserControls.TestClassSelectionFormCaptionExport,
                                                        UserControls.SelectListSelectAllCheckboxCaption,
                                                        UserControls.TestClassSelectionFormCommitTestExport,
                                                        true, UserControls.TestClassSelectionFormOptionalCheckBoxTextExport);
                var list = TestClassManager.GetTestModulesFromVBProject();

                foreach (var item in list)
                {
                    dataContext.Items.Add(new CheckableItem(item.Name));
                }

                var form = new ImportExportWindow(dataContext);
                dataContext.RefreshList += (sender, e) =>
                {
                    e.Items.Clear();
                    foreach (var item in TestClassManager.GetTestModulesFromVBProject())
                    {
                        e.Items.Add(new CheckableItem(item.Name));
                    }
                };
                dataContext.ItemsSelected += (sender, e) =>
                {
                    var sb = new StringBuilder();
                    bool errRaised = false;
                    foreach (var item in e.Items)
                    {
                        try
                        {
                            if (item.IsChecked)
                            {
                                TestClassManager.ExportTestClass(item.Name);
                                item.IsChecked = false;
                                sb.AppendLine("  - " + item.Name);
                            }
                        }
                        catch (Exception ex)
                        {
                            errRaised = true;
                            UITools.ShowException(ex);
                        }
                    }

                    var msg = string.Format(MessageStrings.TestExportedCommitMessage, sb.ToString());
                    UITools.ShowMessage(msg);

                    if (!errRaised)
                    {
                        form.Close();
                    }
                };
                dataContext.PropertyChanged += (sender, e) =>
                {
                    if (e.PropertyName == nameof(dataContext.OptionalCheckboxChecked))
                    {
                        dataContext.CommitButtonText = dataContext.OptionalCheckboxChecked
                            ? UserControls.TestClassSelectionFormCommitTestExportAndRemove
                            : UserControls.TestClassSelectionFormCommitTestExport;
                    }
                };

                SetDialogPosition(form);
                form.ShowDialog();

                dataContext.RefreshList -= (sender, e) => { };
                dataContext.ItemsSelected -= (sender, e) => { };
                dataContext.PropertyChanged -= (sender, e) => { };
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
            }
        }

        private void SetDialogPosition(System.Windows.Window dialog)
        {
            var scaleFactor = UITools.GetScalingFactor();
            var width = dialog.Width;
            var height = dialog.Height;

            var mainWindow = TestClassManager.ActiveVBProject.VBE.MainWindow;

            dialog.Top = (mainWindow.Top + mainWindow.Height / 2) / scaleFactor - height / 2;
            dialog.Left = (mainWindow.Left + mainWindow.Width / 2) / scaleFactor - width / 2;
        }

        #region ICommandBarsAdapterClient support

        public void SubscribeToCommandBarAdapter(VbeCommandBarAdapter commandBarAdapter)
        {
            using (new BlockLogger())
            {
                if (!(commandBarAdapter is AccUnitCommandBarAdapter accUnitCommandBarAdapter)) return;

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
