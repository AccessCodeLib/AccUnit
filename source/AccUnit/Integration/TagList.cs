using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit
{
    public class TagList : List<ITestItemTag>, ITagList
    {
        public TagList()
        {
        }

        public TagList(IEnumerable<ITestItemTag> tags)
        {
            AddRange(tags);
        }

        public new void AddRange(IEnumerable<ITestItemTag> tags)
        {
            foreach (var tag in
                from tag in tags
                let match = Find(x => x.Name == tag.Name)
                where match is null
                select tag)
            {
                Add(tag);
            }
        }

        public bool IsMatch(IEnumerable<ITestItemTag> tags)
        {
            return Count != 0 && tags.Any(tag => Find(x => x.Name == tag.Name) != null);
        }
    }

}
