using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("1CF26408-E864-4A70-A2B3-4423AA410A1F")]
    public interface ITestResultReporter
    {
        ITestResultCollector TestResultCollector { get; set; }
    }
}
