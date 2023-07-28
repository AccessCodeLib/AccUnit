using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using Microsoft.VisualBasic;
using System.Collections.Generic;
using System.Linq;

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
            if (_testClassInstance is null)
            {
                return;
            }

            var vbc = vbProject.VBComponents.Item(Name);
            var codeReader = new CodeModuleReader(vbc.CodeModule);
            var members = codeReader.Members.FindAll(true).FindAll(m => m.ProcKind == vbext_ProcKind.vbext_pk_Proc);

            foreach (var member in members)
            {
                var fixtureMember = GetTestFixtureMember(vbProject, Name, member.Name);
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

        public static ITestFixtureMember GetTestFixtureMember(VBProject vbProject, string fixtureName, string memberName)
        {
            var fixtureMember = new TestFixtureMember(memberName);

            var testClassReader = new TestClassReader(vbProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            return fixtureMember;
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
                return _tests;
            }
        }

        public void FillTestListFromTestClassInstance(VBProject vbProject)
        {
            foreach (var member in from member in _fixtureMembers.Tests
                                   select member)
            {
                _tests.Add(CreateTest(vbProject, this, member.Name));
            }
        }

        public static ITest CreateTest(VBProject vbProject, ITestFixture testFixture, string testMethodName)
        {
            var memberInfo = TestFixture.GetTestFixtureMember(vbProject, testFixture.Name, testMethodName).TestClassMemberInfo;

            if (memberInfo.TestRows.Count > 0)
            {
                return new RowTest(testFixture, memberInfo);
            }

            var test = new MethodTest(testFixture, memberInfo);
            return test;
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