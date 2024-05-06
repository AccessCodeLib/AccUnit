using System;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface IVBATestSuite : ITestSuite, IDisposable
    {
        IVBATestSuite Add(object testToAdd);
        IVBATestSuite AddByClassName(string className);
        IVBATestSuite AddFromVBProject();
        TestClassMemberInfo GetTestClassMemberInfo(string classname, string membername);
    }
}
