using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Integration
{
    internal class ParamTest : BaseTest, IParamTest
    {
        public ParamTest(ITestFixture fixture, ITestClassMemberInfo testClassMemberInfo, string testRowId, IEnumerable<object> parameters)
            : base(fixture, testClassMemberInfo)
        {
            Parameters = parameters;
        }

        public IEnumerable<object> Parameters { get; private set; }
    }
}
