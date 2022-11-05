using System.Collections.Generic;
using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit
{
    public interface ITestClassMemberInfo
    {
        string Name { get; }
        TestClassInfo Parent { get; }
        IgnoreInfo IgnoreInfo { get; }
        IList<int> TestRowFilter { get; }
        TagList Tags { get; }
        bool DoAutoRollback { get; }
        bool IsMatch(IEnumerable<TestItemTag> tags);
        List<ITestRow> TestRows { get; }
        IList<VbMsgBoxResult> MsgBoxResults { get; }
        string ShowAs { get; }
        string DisplayName { get; }
    }

}
