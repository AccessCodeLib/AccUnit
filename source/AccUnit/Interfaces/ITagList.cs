using System.Collections.Generic;

namespace AccessCodeLib.AccUnit
{
    public interface ITagList : IEnumerable<ITestItemTag>
    {
        void AddRange(IEnumerable<ITestItemTag> tags);
        bool IsMatch(IEnumerable<ITestItemTag> tags);
    }
}