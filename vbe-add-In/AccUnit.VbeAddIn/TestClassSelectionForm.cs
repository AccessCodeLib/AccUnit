using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using AccessCodeLib.Common.VBIDETools;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public sealed partial class TestClassSelectionForm : Form
    {
        public enum SelectionMode { Export, Import };

        public delegate void CommitSelectedTestsEventHandler(TestClassSelectionForm sender, TestComponentsEventArgs e, ref bool close);
        public event CommitSelectedTestsEventHandler TestsSelected;

        public delegate void RefreshTestListEventHandler(TestClassSelectionForm sender, TestComponentsEventArgs e);
        public event RefreshTestListEventHandler RefreshTestList;

        private TestClassSelectionForm()
        {
            InitializeComponent();
            testListUserControl.TestsSelected += TestListUserControlTestsSelected;
            testListUserControl.RefreshTestList += TestListUserControlRefreshTestList;
        }

        public TestClassSelectionForm(SelectionMode mode, IEnumerable<CodeModuleInfo> list = null)
            : this()
        {
            CurrentMode = mode;
            switch (mode)
            {
                case SelectionMode.Export:
                    Text = Resources.UserControls.TestClassSelectionFormCaptionExport;
                    Icon = UITools.ConvertBitmapToIcon(Resources.Icons.MoveToFolder);
                    testListUserControl.SetCommitButtonLayout(Resources.UserControls.TestClassSelectionFormCommitTestExport, Resources.Icons.MoveToFolder);    
                    break;
                case SelectionMode.Import:
                    Text = Resources.UserControls.TestClassSelectionFormCaptionImport;
                    Icon = UITools.ConvertBitmapToIcon(Resources.Icons.ImportFromFolder);
                    testListUserControl.SetCommitButtonLayout(Resources.UserControls.TestClassSelectionFormCommitTestImport, Resources.Icons.ImportFromFolder);
                    break;
                default:
                    throw new NotImplementedException();
            }

            if (list != null)
            {
                Add(list);
            }
        }

        public SelectionMode CurrentMode { get; private set; }

        public TestClassSelectionForm(string commitbuttonText, Image commitbuttonImage)
            : this()
        {
            testListUserControl.SetCommitButtonLayout(commitbuttonText, commitbuttonImage);
        }

        void TestListUserControlRefreshTestList(object sender, TestComponentsEventArgs e)
        {
            if (RefreshTestList != null)
            {
                RefreshTestList(this, e);
            }
        }

        void TestListUserControlTestsSelected(object sender, TestComponentsEventArgs e)
        {
            if (TestsSelected == null) return;
            var close = false;
            TestsSelected(this, e, ref close);
            if (close)
            {
                Close();
            }
        }

        public void Add(IEnumerable<CodeModuleInfo> testclasses)
        {
            testListUserControl.Add(testclasses);
        }
    }
}
