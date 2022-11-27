using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("823F42EC-DF58-4251-8B57-437D2D5D4BF9")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITestResultSummary : ITestResult
    {
        int Count { get; }
        
        [DispId(0)]
        ITestResult Item(int Index);
    }

}
