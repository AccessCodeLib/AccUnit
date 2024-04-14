using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.Common;
using AccessCodeLib.Common.VBIDETools;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public partial class TestListControl : SelectListControl
    {
        //public delegate void CommitSelectedTestsEventHandler(TestClassList list);
        //public event CommitSelectedTestsEventHandler TestsSelected;
        public event EventHandler<TestComponentsEventArgs> TestsSelected;

        //public delegate void RefreshTestListEventHandler(ref TestClassList list);
        //public event RefreshTestListEventHandler RefreshTestList;
        public event EventHandler<TestComponentsEventArgs> RefreshTestList;

        public void Add(IEnumerable<CodeModuleInfo> modules)
        {
            base.Add(modules.ToArray());
        }

        protected override void RaiseCommit()
        {
            if (TestsSelected == null) return;
            var list = new List<CodeModuleInfo>((from CodeModuleInfo c in itemListBox.CheckedItems
                                                   select c).ToList());
            TestsSelected(this, new TestComponentsEventArgs(list));
        }

        public IList<TestClassInfo> CheckedItems
        {
            get
            {
                return (from TestClassInfo c in itemListBox.CheckedItems
                        select c).ToList();
            }
        }

        public override void RefreshList()
        {
            try
            {
                IEnumerable<CodeModuleInfo> list = null;
                if (RefreshTestList != null)
                {
                    var e = new TestComponentsEventArgs();
                    RefreshTestList(this, e);
                    list = e.Components;
                }

                List<string> checkedNames = null;

                if (itemListBox.CheckedItems.Count > 0)
                {
                    checkedNames = new List<string>(from TestClassInfo info in itemListBox.CheckedItems
                                                    select info.Name);
                }

                itemListBox.Items.Clear();
                if (list == null)
                    return;

                Add(list);

                var classNames = new List<string>((from TestClassInfo info in itemListBox.Items
                                                   select info.Name));

                RefreshSelectedItems(classNames, checkedNames);
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
            }

        }

    }

    internal static class TestListUserControlInfo
    {
        public const string ProgID = @"AccUnit.GUI.TestListControl";
        public const string PositionGuid = @"F05AAD71-D046-45C6-927F-48D326557518";
        public const string ClassName = @"AccessCodeLib.AccUnit.AddIn.TestListControl";
    }
}
