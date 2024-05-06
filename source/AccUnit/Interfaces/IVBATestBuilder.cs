using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface IVBATestBuilder : ITestBuilder
    {
        //VBProject ActiveVBProject { get; }
        //object HostApplication { get; }
    }

    public interface ITestBuilder : IDisposable
    {
        bool TestToolsActivated { get; }
        object CreateObject(string className);
        object CreateTest(object testToAdd, ITestClassMemberList memberFilter);
        object CreateTest(string className);
        IEnumerable<object> CreateTests(IEnumerable<TestClassInfo> testClasses);
        IEnumerable<object> CreateTestsFromVBProject();
        void DeleteFactoryCodeModule();
        void RefreshFactoryCodeModule();
    }

}