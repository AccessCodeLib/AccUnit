using AccessCodeLib.AccUnit.VbeAddIn.TestExplorer;
using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableTestItemsEventArgs : EventArgs
    {
        public CheckableTestItemsEventArgs()
        {
        }

        public CheckableTestItemsEventArgs(TestItems items)
        {
            Items = items;
        }

        public TestItems Items { get; set; }
    }

    public class CheckableItemsEventArgs : EventArgs
    {
        public CheckableItemsEventArgs()
        {
        }

        public CheckableItemsEventArgs(ICollection<ICheckableItem> items)
        {
            Items = items;
        }

        public ICollection<ICheckableItem> Items { get; set; }
    }
}