using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.Integration
{
    internal abstract class BaseTest : ITest
    {
        public BaseTest(ITestFixture fixture, ITestClassMemberInfo testClassMemberInfo)
        {
            Fixture = fixture;
            Name = testClassMemberInfo.Name;
            MethodName = testClassMemberInfo.Name;
            TestClassMemberInfo = testClassMemberInfo;
        }

        protected virtual string FormattedFullName()
        {
            return $"{Fixture.Name}.{MethodName}";
        }

        public ITestFixture Fixture { get; private set; }

        public string MethodName { get; private set; }

        public string FullName { get { return FormattedFullName(); } }

        public string DisplayName { get; set; }
        public RunState RunState { get; set; }

        public string Name { get; private set; }

        public ITestClassMemberInfo TestClassMemberInfo { get; private set; }
    }
}
