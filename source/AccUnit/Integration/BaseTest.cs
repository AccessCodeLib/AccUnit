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
            
            SetFullName();
        }

        protected virtual void SetFullName()
        {
            FullName = $"{Fixture.Name}.{MethodName}";
        }

        public ITestFixture Fixture { get; private set; }

        public string MethodName { get; private set; }

        public string FullName { get; private set; }

        public string DisplayName { get; set; }
        public RunState RunState { get; set; }

        public string Name { get; private set; }

        public ITestClassMemberInfo TestClassMemberInfo { get; private set; }
    }
}
