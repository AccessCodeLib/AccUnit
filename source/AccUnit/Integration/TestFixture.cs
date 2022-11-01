using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.AccUnit.Interop;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.VbaProjectManagement;
using System.Collections.Generic;
using System.Linq;
using TLI;

namespace AccessCodeLib.AccUnit
{

    internal class TestFixture : ITestFixture
    {
        private readonly object _testClassInstance;
        private readonly IList<ITest> _tests = new List<ITest>();
        private readonly TestFixtureMembers _fixtureMembers = new TestFixtureMembers();
        
        public TestFixture(object testClassInstance)
        {
            _testClassInstance = testClassInstance;
            Name = AccessCodeLib.Common.VBIDETools.TypeLib.TypeLibTools.GetTLIInterfaceInfoName(testClassInstance);
            FullName = Name;
            FillInstanceMembers();
        }
        
        public TestFixture(string name, object testToAdd)
        {
            _testClassInstance = testToAdd;
            Name = name;
            FullName = name;
            FillInstanceMembers();
        }

        public ITestFixtureMembers Members
        {
            get
            {
                return _fixtureMembers;
            }
        }

        private void FillInstanceMembers()
        {
            if (_testClassInstance == null)
            {
                return;
            }

            TLI.Members members = AccessCodeLib.Common.VBIDETools.TypeLib.TypeLibTools.GetTLIInterfaceMembers(_testClassInstance);
            foreach(TLI.MemberInfo member in members)
            {
                var fixtureMember = new TestFixtureMember(member.Name);
                _fixtureMembers.Add(fixtureMember);
                
                if (fixtureMember.IsFixtureSetup)
                {
                    HasFixtureSetup = true;
                }
                else if (fixtureMember.IsSetup)
                {
                    HasSetup = true;
                }
                else if (fixtureMember.IsTeardown)
                {
                    HasTeardown = true;
                }
                else if (fixtureMember.IsFixtureTeardown)
                {
                    HasFixtureTeardown = true;
                }
            }
        }

        public object Instance
        {
            get { return _testClassInstance; }
        }

        public string Name { get; private set; }

        public string FullName { get; set; }

        public RunState RunState { get; set; }

        public IEnumerable<ITest> Tests
        {
            get 
            {
                if (_tests.Count == 0)
                {
                    FillTestListFromTestClassInstance();
                }
                return _tests; 
            }
        }

        private void FillTestListFromTestClassInstance()
        {
            foreach (var member in from member in _fixtureMembers.Tests
                                   select member)
            {
                _tests.Add(new MethodTest(this, member.Name));
            }
        }
        
        public bool HasFixtureSetup { get; set; }
        public bool HasSetup { get; set; }
        public bool HasTeardown { get; set; }
        public bool HasFixtureTeardown { get; set; }

        public void Add(ITest test)
        {
            _tests.Add(test);
        }
    }
}