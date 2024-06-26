﻿using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Integration
{
    public class RowTest : IRowTest
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
                if (row.Name is null)
                {
                    row.Name = i.ToString();
                }
                var paramTestClassMemberInfo = new TestClassMemberInfo(TestClassMemberInfo, row.IgnoreInfo, row.Tags);
                var paramTest = new ParamTest(Fixture, this, paramTestClassMemberInfo, row.Name, row.Args);
                paramTests.Add(paramTest);
            }

            ParamTests = paramTests;

        }

        public ITestFixture Fixture { get; private set; }

        public string MethodName { get; private set; }

        public string FullName { get; private set; }

        protected string _displayName;
        public string DisplayName { get { return _displayName ?? Name; } set { _displayName = value; } }
        public RunState RunState { get; set; }

        public string Name { get; private set; }

        public ITestClassMemberInfo TestClassMemberInfo { get; private set; }

        public IEnumerable<ITestRow> Rows { get; private set; }

        public IEnumerable<IParamTest> ParamTests { get; private set; }

        public object Parent => Fixture;
    }
}
