using AccessCodeLib.AccUnit.VbeAddIn.TestExplorer;
using System;

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
}