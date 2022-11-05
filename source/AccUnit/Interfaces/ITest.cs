namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITest : ITestData
    {
        ITestFixture Fixture { get; } 
        string MethodName { get; }
        string DisplayName { get; set; }
        RunState RunState { set; get; }
        ITestClassMemberInfo TestClassMemberInfo { get; }
    }
}
