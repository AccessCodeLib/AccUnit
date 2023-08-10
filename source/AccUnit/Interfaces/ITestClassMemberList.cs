using System.Collections.Generic;

namespace AccessCodeLib.AccUnit
{
    public interface ITestClassMemberList : IEnumerable<TestClassMemberInfo>
    {
        ITagList Tags { get; }
        ITestClassMemberList Filter(IEnumerable<ITestItemTag> tags);
    }
}