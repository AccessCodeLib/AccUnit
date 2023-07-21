using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit
{
    
    public class TestFixtureMembers : List<ITestFixtureMember>, ITestFixtureMembers
    {
        public ITestFixtureMember FixtureSetup
        {
            get
            {
                return Find(m => m.IsSetup);
            }
        }
        
        public ITestFixtureMember Setup
        {
            get
            {
                return Find(m => m.IsSetup);
            }
        }
        
        public ITestFixtureMember Teardown
        {
            get
            {
                return Find(m => m.IsTeardown);
            }
        }

        public ITestFixtureMember FixtureTeardown
        {
            get
            {
                return Find(m => m.IsFixtureTeardown);
            }
        }

        public IEnumerable<ITestFixtureMember> Tests
        {
            get
            {
                return FindAll(m => m.IsTest);
            }
        }

    }

    public interface ITestFixtureMember
    {
        string Name { get; }
        bool IsFixtureSetup { get; }
        bool IsSetup { get; }
        bool IsTeardown { get; }
        bool IsFixtureTeardown { get; }
        bool IsTest { get; }
        TestClassMemberInfo TestClassMemberInfo { get; set; }
    }

    public class TestFixtureMember : ITestFixtureMember
    {
        public TestFixtureMember (string name)
        {
            Name = name;
            SetType();
        }

        private void SetType()
        {
            if (Name.Equals("FixtureSetup", System.StringComparison.InvariantCultureIgnoreCase))
            {
                IsFixtureSetup = true;
            }
            else if (Name.Equals("Setup", System.StringComparison.InvariantCultureIgnoreCase))
            {
                IsSetup = true;
            }
            else if (Name.Equals("Teardown", System.StringComparison.InvariantCultureIgnoreCase))
            {
                IsTeardown = true;
            }
            else if (Name.Equals("FixtureTeardown", System.StringComparison.InvariantCultureIgnoreCase))
            {
                IsFixtureTeardown = true;
            }
            else
            {
                IsTest = true;
            }
        }

        public string Name { get; private set; }
        public bool IsFixtureSetup { get; private set; }
        public bool IsSetup { get; private set; }
        public bool IsTeardown { get; private set; }
        public bool IsFixtureTeardown { get; private set; }
        public bool IsTest { get; private set; }

        public TestClassMemberInfo TestClassMemberInfo { get; set; }

    }
}