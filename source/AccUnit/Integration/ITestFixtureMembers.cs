using System.Collections.Generic;

namespace AccessCodeLib.AccUnit
{
    public interface ITestFixtureMembers //: IList<TestFixtureMember>
    {
        ITestFixtureMember FixtureSetup { get; }
        ITestFixtureMember Setup { get; }
        ITestFixtureMember Teardown { get; }
        ITestFixtureMember FixtureTeardown { get; }
        IEnumerable<ITestFixtureMember> Tests { get; }
    }
}