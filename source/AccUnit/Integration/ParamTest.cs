using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Integration
{
    internal class ParamTest : BaseTest, IParamTest
    {
        private readonly string _testRowId = string.Empty;

        public ParamTest(ITestFixture fixture, ITestClassMemberInfo testClassMemberInfo, string testRowId, IEnumerable<object> parameters)
            : base(fixture, testClassMemberInfo)
        {
            _testRowId = testRowId;
            Parameters = parameters;
        }

        protected override string FormattedFullName()
        {
            return $"{Fixture.Name}.{MethodName}.{_testRowId}";
        }

        public IEnumerable<object> Parameters { get; private set; }
    }
}
