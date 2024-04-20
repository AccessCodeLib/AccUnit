using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
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