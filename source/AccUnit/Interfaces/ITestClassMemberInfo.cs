using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.Generic;
using static AccessCodeLib.AccUnit.TestClassMemberInfo;

namespace AccessCodeLib.AccUnit
{
    public interface ITestClassMemberInfo
    {
        string Name { get; }
        TestClassInfo Parent { get; }
        IgnoreInfo IgnoreInfo { get; }
        IList<int> TestRowFilter { get; }
        ITagList Tags { get; }
        bool DoAutoRollback { get; }
        bool IsMatch(IEnumerable<ITestItemTag> tags);
        List<ITestRow> TestRows { get; }
        IList<VbMsgBoxResult> MsgBoxResults { get; }
        string ShowAs { get; }
        string DisplayName { get; }
    }

}
