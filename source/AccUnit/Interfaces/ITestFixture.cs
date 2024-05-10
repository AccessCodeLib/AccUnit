using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("5B540C56-B19B-4A44-BF98-E66BB07C4EB9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITestFixture : ITestData
    {
        object Instance { get; }

        [ComVisible(false)]
        RunState RunState { set; get; }

        [ComVisible(false)]
        ITestFixtureMembers Members { get; }

        [ComVisible(false)]
        IEnumerable<ITest> Tests { get; }   // Definition of the test methods

        [ComVisible(false)]
        IEnumerable<ITestItemTag> Tags { get; }

        [ComVisible(false)]
        bool HasFixtureSetup { get; set; }

        [ComVisible(false)]
        bool HasSetup { get; set; }

        [ComVisible(false)]
        bool HasTeardown { get; set; }

        [ComVisible(false)]
        bool HasFixtureTeardown { get; set; }
    }
}
