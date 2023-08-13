using System.Collections.Generic;

namespace AccessCodeLib.AccUnit
{
    public class TestClassMemberList : List<TestClassMemberInfo>, ITestClassMemberList
    {

        public ITestClassMemberList Filter(IEnumerable<ITestItemTag> tags)
        {
            var list = new TestClassMemberList();
            list.AddRange(FindAll(x => x.IsMatch(tags)));
            return list;
        }

        public ITagList Tags
        {
            get
            {
                var tags = new TagList();
                foreach (var member in this)
                {
                    tags.AddRange(member.Tags);
                }
                return tags;
            }
        }

    }
}
