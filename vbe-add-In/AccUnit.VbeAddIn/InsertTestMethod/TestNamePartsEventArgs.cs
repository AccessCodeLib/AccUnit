using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class TestNamePartsEventArgs : EventArgs
    {
        public TestNamePartsEventArgs()
        {
        }

        public TestNamePartsEventArgs(ICollection<ITestNamePart> items)
        {
            Items = items;
        }

        public ICollection<ITestNamePart> Items { get; set; }
    }

}