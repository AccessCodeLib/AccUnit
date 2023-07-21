﻿namespace AccessCodeLib.AccUnit
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
            if (obj is null)
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
    }
}
