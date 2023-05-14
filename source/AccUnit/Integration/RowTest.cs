using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Integration
{
    internal class RowTest : IRowTest
    {
        public RowTest(ITestFixture fixture, TestClassMemberInfo testClassMemberInfo)
        {
            Fixture = fixture;
            Name = testClassMemberInfo.Name;
            MethodName = testClassMemberInfo.Name;
            FullName = $"{fixture.Name}.{MethodName}";
            TestClassMemberInfo = testClassMemberInfo;
            
            FillRows();
        }

        private void FillRows()
        {
            Rows = TestClassMemberInfo.TestRows;
            var paramTests = new List<IParamTest>();
            int i = 0;
            
            foreach (var row in Rows)
            {
                i++;
                if (row.Name == null)
                {
                    row.Name = i.ToString();
                }
                var paramTest = new ParamTest(Fixture, TestClassMemberInfo, row.Name, row.Args);
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

        public ITestClassMemberInfo TestClassMemberInfo { get; private set; }

        public IEnumerable<ITestRow> Rows { get; private set; }

        public IEnumerable<IParamTest> ParamTests { get; private set; }
    }
}
