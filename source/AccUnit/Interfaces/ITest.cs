using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("96AEE906-564B-4A39-B85C-E47F275CFD51")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITest : ITestData
    {
        ITestFixture Fixture { get; }
        string MethodName { get; }
        string DisplayName { get; set; }
        RunState RunState { set; get; }

        [ComVisible(false)]
        ITestClassMemberInfo TestClassMemberInfo { get; }
    }
}
