using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.AccUnit.Interop;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.VbaProjectManagement;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic;
using Microsoft.Vbe.Interop;

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
            Name = Information.TypeName(testClassInstance);
            FullName = Name;
        }

        public TestFixture(string name, object testToAdd)
        {
            Name = name;
            FullName = name;
            _testClassInstance = testToAdd;
        }

        public ITestFixtureMembers Members
        {
            get
            {
                return _fixtureMembers;
            }
        }

        public void FillInstanceMembers(VBProject vbProject)
        {
            if (_testClassInstance == null)
            {
                return;
            }

            var vbc = vbProject.VBComponents.Item(Name);
            var codeReader = new CodeModuleReader(vbc.CodeModule);
            var members = codeReader.Members.FindAll(true).FindAll(m => m.ProcKind == vbext_ProcKind.vbext_pk_Proc);
            
            foreach(var member in members)
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