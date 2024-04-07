using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.Integration
{
    public abstract class BaseTest : ITest
    {
        public BaseTest(ITestFixture fixture, ITestClassMemberInfo testClassMemberInfo)
        {
            Fixture = fixture;
            Parent = fixture;
            Name = testClassMemberInfo.Name;
            MethodName = testClassMemberInfo.Name;
            TestClassMemberInfo = testClassMemberInfo;
        }

        public BaseTest(ITestFixture fixture, object parentTest, ITestClassMemberInfo testClassMemberInfo)
        {
            Fixture = fixture;
            Name = testClassMemberInfo.Name;
            MethodName = testClassMemberInfo.Name;
            TestClassMemberInfo = testClassMemberInfo;
            Parent = parentTest;
        }

        protected virtual string FormattedFullName()
        {
            return $"{Fixture.Name}.{MethodName}";
        }

        public ITestFixture Fixture { get; private set; }

        public object Parent { get; private set; }

        public string MethodName { get; private set; }

        public string FullName { get { return FormattedFullName(); } }

        protected string _displayName;
        public string DisplayName { get { return _displayName?? Name; } set { _displayName = value; } }
        public RunState RunState { get; set; }

        public string Name { get; private set; }

        public ITestClassMemberInfo TestClassMemberInfo { get; private set; }
    }
}
