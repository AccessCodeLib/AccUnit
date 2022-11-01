using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("14576D04-E4A4-482F-931C-C46C5F39F294")]
    public interface ITestManagerBridge
    {
        void InitTestManager(TestManager TestManager);
        TestManager GetTestManager();
    }
}
