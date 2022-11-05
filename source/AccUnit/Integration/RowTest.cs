using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace AccessCodeLib.AccUnit.Integration
{
    internal class RowTest : IRowTest
    {
        private TestClassMemberInfo _testClassMemberInfo;
        
        public RowTest(ITestFixture fixture, TestClassMemberInfo testClassMemberInfo)
        {
            Fixture = fixture;
            Name = testClassMemberInfo.Name;
            MethodName = testClassMemberInfo.Name;
            FullName = $"{fixture.Name}.{MethodName}";

            _testClassMemberInfo = testClassMemberInfo;
            FillRows();
        }

        private void FillRows()
        {
            Rows = _testClassMemberInfo.TestRows;
            var paramTests = new List<IParamTest>();
            
            foreach (var row in Rows)
            {
                var paramTest = new ParamTest(Fixture, MethodName, row.Name, row.Args);
                paramTests.Add(paramTest);
            }

            ParamTests = paramTests;

        }

        public ITestFixture Fixture { get; private set; }

        public string MethodName { get; private set; }

        public string FullName { get; private set; }

        public string DisplayName { get; set; }
        public RunState RunState { get; set; }

        public string Name { get; private set; }

        public IEnumerable<ITestRow> Rows { get; private set; }

        public IEnumerable<IParamTest> ParamTests { get; private set; }
    }
}
