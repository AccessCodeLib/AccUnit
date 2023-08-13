using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITestRow
    {
        int Index { get; set; }
        IList<object> Args { get; }
        string Name { get; set; }
        ITagList Tags { get; }
        ITestRow SetName(string name);
        ITestRow ClickingMsgBox(params VbMsgBoxResult[] args);
        ITestMessageBox TestMessageBox { get; set; }
        ITestRow Ignore(string comment = "");
        IgnoreInfo IgnoreInfo { get; }
        string TestFixtureRowName { get; }
    }

    /*
    public interface ITestRow
    {
        int Index { get; set; }
        IList<object> Args { get; }
        string Name { get; set; }
        ITestRow ClickingMsgBox(params VBA.VbMsgBoxResult[] args);
        TestMessageBox TestMessageBox { get; set; }
    }
    */
    [ComVisible(true)]
    [Guid("B9BA0F7E-5FEB-4FCF-BD6C-7C6F33A1324E")]
#pragma warning disable IDE1006 // Benennungsstile
    public interface _ITestRow
#pragma warning restore IDE1006 // Benennungsstile
    {
        string Name { get; set; }
        ITestRow ClickingMsgBox(params VbMsgBoxResult[] args);
    }
}
