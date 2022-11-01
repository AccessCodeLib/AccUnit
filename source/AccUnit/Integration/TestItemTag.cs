namespace AccessCodeLib.AccUnit
{
    public class TestItemTag
    {

        public TestItemTag(string name)
        {
            Name = name;
        }

        public string Name { get; private set; }

        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
                return false;
            if (ReferenceEquals(this, obj))
                return true;
            if (obj.GetType() != typeof(TestItemTag))
                return false;
            var other = (TestItemTag)obj;
            return Equals(other.Name, Name);
        }

        public override int GetHashCode()
        {
            return Name.GetHashCode();
        }

        /*
        public void AddTestClassMemberInfo(TestClassMemberInfo m)
        {
            _testClassMembers.Add(m);
            if (m.Parent != null && _testClasses.Contains(m.Parent) == false)
            {
                _testClasses.Add(m.Parent);
            }
        }

        public void AddTestClassInfo(TestClassInfo c)
        {
            _testClasses.Add(c);
            foreach (TestClassMemberInfo m in c.Members)
            {
                if (_testClassMembers.Contains(m) == false)
                {
                    _testClassMembers.Add(m);
                }
            }
        }

        public void Merge(TestItemTag tag)
        {
            foreach (TestClassInfo c in tag.TestClasses)
            {
                AddTestClassInfo(c);
            }
            foreach (TestClassMemberInfo m in tag.TestClassMembers)
            {
                AddTestClassMemberInfo(m);
            }
        }

        List<TestClassMemberInfo> _testClassMembers = new List<TestClassMemberInfo>();
        public List<TestClassMemberInfo> TestClassMembers { get { return _testClassMembers; } }

        List<TestClassInfo> _testClasses = new List<TestClassInfo>();
        public List<TestClassInfo> TestClasses { get { return _testClasses; } }
        */

    }
}
