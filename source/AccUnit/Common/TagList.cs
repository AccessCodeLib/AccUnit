using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit.Common
{
    public class TagList : List<TestItemTag>
    {
        public TagList()
        {
        }

        public TagList(IEnumerable<TestItemTag> tags)
        {
            AddRange(tags);
        }

        public new void AddRange(IEnumerable<TestItemTag> tags)
        {
            foreach (var newTag in tags.Where(t => !Contains(t)))
            {
                Add(newTag);
            }
        }

        public bool IsMatch(IEnumerable<TestItemTag> tags)
        {
            return this.Intersect(tags).Any();
        }

        public override string ToString()
        {
            return string.Join(", ", this.Select(tit => tit.Name).ToArray());
        }
    }

}
