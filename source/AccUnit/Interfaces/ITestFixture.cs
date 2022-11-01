using AccessCodeLib.AccUnit.Interop;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITestFixture : ITestData
    {
        object Instance { get; }
        RunState RunState { set; get; }
        ITestFixtureMembers Members { get; }
        IEnumerable<ITest> Tests { get; }   // Definition der Testmethoden

        bool HasFixtureSetup { get; set; }
        bool HasSetup { get; set; }
        bool HasTeardown { get; set; }
        bool HasFixtureTeardown { get; set; }
    }
}
