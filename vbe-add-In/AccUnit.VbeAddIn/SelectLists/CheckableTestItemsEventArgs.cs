using AccessCodeLib.AccUnit.VbeAddIn.TestExplorer;
using System;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableTestItemsEventArgs : EventArgs
    {
        public CheckableTestItemsEventArgs()
        {
        }

        public CheckableTestItemsEventArgs(CheckableItems<TestItem> items)
        {
            Items = items;
        }

        public CheckableItems<TestItem> Items { get; set; }
    }
}