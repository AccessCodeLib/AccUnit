using System.Collections.Generic;

namespace AccessCodeLib.AccUnit
{
    public class TestClassList : List<TestClassInfo>
    {
        public TestClassList()
        {
        }

        public TestClassList(IEnumerable<TestClassInfo> testclassinfoList)
        {
            AddRange(testclassinfoList);
        }

        public new void AddRange(IEnumerable<TestClassInfo> collection)
        {
            base.AddRange(collection);
            if (_tags != null)
                AddTags(collection);
        }

        private TagList _tags;
        public TagList Tags
        {
            get
            {
                if (_tags is null)
                    FillTagList();
                return _tags;
            }
        }

        private void FillTagList()
        {
            _tags = new TagList();
            AddTags(this);
        }

        private void AddTags(IEnumerable<TestClassInfo> collection)
        {
            foreach (var testclass in collection)
            {
                _tags.AddRange(testclass.Tags);
            }
        }
    }

}
