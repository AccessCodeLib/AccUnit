using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.Integration
{
    internal class MethodTest : ITest
    {
        public MethodTest(ITestFixture fixture, string methodName)
        {
            Fixture = fixture;
            Name = methodName;
            MethodName = methodName;
            FullName = $"{fixture.Name}.{methodName}";
        }

        public ITestFixture Fixture { get; private set; }

        public string MethodName { get; private set; }

        public string FullName { get; private set; }

        public string DisplayName { get; set; }
        public RunState RunState { get; set; }

        public string Name { get; private set; }
    }
}
