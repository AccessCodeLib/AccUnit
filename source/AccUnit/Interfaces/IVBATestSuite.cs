namespace AccessCodeLib.AccUnit.Interfaces
{
    
    public interface IVBATestSuite : ITestSuite
    {
        IVBATestSuite       AddByClassName(string className);
        IVBATestSuite       AddFromVBProject();
        void                Dispose();
        TestClassMemberInfo GetTestClassMemberInfo(string classname, string membername);
    }
}
