using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Common
{
    public class TestClassMemberList : List<TestClassMemberInfo>
    {

        public TestClassMemberList Filter(TagList tags)
        {
            var list = new TestClassMemberList();
            list.AddRange(FindAll(x => (x.IsMatch(tags))));
            return list;
        }

        public TagList Tags
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
