using AccessCodeLib.Common.VBIDETools;
using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class TestComponentsEventArgs : EventArgs
    {
        public TestComponentsEventArgs()
        {
        }

        public TestComponentsEventArgs(IEnumerable<CodeModuleInfo> components)
        {
            Components = components;
        }

        public IEnumerable<CodeModuleInfo> Components { get; set; }
    }
}