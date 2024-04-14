using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Integration
{
    public class ParamTest : BaseTest, IParamTest, IRowTestId
    {
        private readonly string _testRowId = string.Empty;

        public ParamTest(ITestFixture fixture, ITestClassMemberInfo testClassMemberInfo, string testRowId, IEnumerable<object> parameters)
            : base(fixture, testClassMemberInfo)
        {
            _testRowId = testRowId;
            Parameters = parameters;
        }

        public ParamTest(ITestFixture fixture, ITest parent, ITestClassMemberInfo testClassMemberInfo, string testRowId, IEnumerable<object> parameters)
            : base(fixture, parent, testClassMemberInfo)
        {
            _testRowId = testRowId;
            Parameters = parameters;
        }

        protected override string FormattedFullName()
        {
            return $"{Fixture.Name}.{MethodName}.{_testRowId}";
        }

        public IEnumerable<object> Parameters { get; private set; }

        public string RowId => _testRowId;
    }

    public interface IRowTestId
    {
        string RowId { get; }   
    }
}
