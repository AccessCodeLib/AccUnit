using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITestRow
    {
        int Index { get; set; }
        IList<object> Args { get; }
        string Name { get; set; }
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
    public interface _ITestRow
    {
        string Name { get; set; }
        ITestRow ClickingMsgBox(params VbMsgBoxResult[] args);
    }
}
