using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.Integration
{
    internal class ParamTest : IParamTest
    {
        public ParamTest(ITestFixture fixture, string methodName, string testRowId, IEnumerable<object> parameters)
        {
            Fixture = fixture;
            Name = $"{methodName}.{testRowId}";
            MethodName = methodName;
            FullName = $"{fixture.Name}.{methodName}.{testRowId}";
            Parameters = parameters;
        }

        public ITestFixture Fixture { get; private set; }

        public string MethodName { get; private set; }

        public string FullName { get; private set; }

        public string DisplayName { get; set; }
        public RunState RunState { get; set; }

        public string Name { get; private set; }

        public IEnumerable<object> Parameters { get; private set; }
    }
}
