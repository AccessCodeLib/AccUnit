using System;
using System.Collections;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    [ComVisible(true)]
    [Guid("EFA775FF-F599-4649-AE24-3614865FA96B")]
    public partial class SelectListControl : UserControl
    {

        public delegate void CommitEventHandler(object sender, ICollection checkedItems);
        public event CommitEventHandler ItemsSelected;

        public delegate void RefreshItemListEventHandler(object sender, ICollection items);
        public event RefreshItemListEventHandler RefreshItemList;

        public IList Items { get { return itemListBox.Items; } }

        public SelectListControl()
        {
            InitializeComponent();
            var refreshButtonToolTip = new ToolTip();
            refreshButtonToolTip.SetToolTip(RefreshButton, Resources.ToolTips.SelectListUserControl_RefreshButton);
        }

        public SelectListControl(string commitbuttonText, Image commitbuttonImage)
            : this()
        {
            SetCommitButtonLayout(commitbuttonText, commitbuttonImage);
        }

        public void SetCommitButtonLayout(string commitbuttonText, Image commitbuttonImage)
        {
            CommitButton.Text = commitbuttonText;
            CommitButton.Image = commitbuttonImage;
        }

        public void Add(object[] items)
        {
            itemListBox.Items.AddRange(items);
            RefreshSelectAllCheckBoxCheckState();
        }

        protected virtual void RaiseCommit()
        {
            if (ItemsSelected != null)
                ItemsSelected(this, itemListBox.CheckedItems);
        }

        private void CommitSelectedTestsButtonClick(object sender, EventArgs e)
        {
            RaiseCommit();
        }

        private void RefreshButtonClick(object sender, EventArgs e)
        {
            RefreshList();
            Refresh();
            itemListBox.Refresh();
        }

        public virtual void RefreshList()
        {
            if (RefreshItemList != null)
                RefreshItemList(this, itemListBox.Items);
        }

        protected void RefreshSelectedItems(IList<string> itemNames, IEnumerable<string> checkedNames)
        {
            if (checkedNames != null)
            {
                foreach (var name in checkedNames)
                {
                    _disableAutoItemCheck = true;
                    var index = itemNames.IndexOf(name);
                    if (index >= 0)
                        itemListBox.SetItemChecked(index, true);
                    _disableAutoItemCheck = false;
                }
            }
            RefreshSelectAllCheckBoxCheckState();
        }

        protected void RefreshSelectAllCheckBoxCheckState(int nextCheck = 0)
        {
            _disableAutoItemCheck = true;
            var checkValue = nextCheck + itemListBox.CheckedItems.Count;

            if (checkValue == 0)
            {
                selectAllCheckBox.CheckState = CheckState.Unchecked;
            }
            else if (checkValue == itemListBox.Items.Count)
            {
                selectAllCheckBox.CheckState = CheckState.Checked;
            }
            else
            {
                selectAllCheckBox.CheckState = CheckState.Indeterminate;
            }
            _disableAutoItemCheck = false;
        }

        private bool _disableAutoItemCheck;

        private void ItemListBoxItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_disableAutoItemCheck)
            {
                return;
            }
            var nextCheck = (itemListBox.GetItemChecked(itemListBox.SelectedIndex)) ? -1 : 1 ;
            RefreshSelectAllCheckBoxCheckState(nextCheck);
        }

        private void SelectAllCheckBoxCheckStateChanged(object sender, EventArgs e)
        {
            if (_disableAutoItemCheck)
            {
                return;
            }

            if (selectAllCheckBox.CheckState == CheckState.Indeterminate)
            {
                selectAllCheckBox.Checked = false;
                return;
            }

            _disableAutoItemCheck = true;

            var checkAll = false;
            if (selectAllCheckBox.Checked)
            {
                checkAll = true;
            }

            for (var i = 0; i < itemListBox.Items.Count; i++)
            {
                if (itemListBox.GetItemChecked(i) != checkAll)
                    itemListBox.SetItemChecked(i, checkAll);
            }

            _disableAutoItemCheck = false;
        }

        private void SelectListControlResize(object sender, EventArgs e)
        {
            Refresh();
        }
    }
}
