using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CommitInsertTestMethodEventArgs : EventArgs
    {
        public CommitInsertTestMethodEventArgs(string methodUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            MethodUnderTest   = methodUnderTest;
            StateUnderTest    = stateUnderTest;
            ExpectedBehaviour = expectedBehaviour;
        }

        public string MethodUnderTest { get; private set; }
        public string StateUnderTest { get; private set; }
        public string ExpectedBehaviour { get; private set; }
    }

    public class CommitInsertTestMethodsEventArgs : EventArgs
    {
        public CommitInsertTestMethodsEventArgs(string testClass, IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            TestClass = testClass;
            MethodsUnderTest = methodsUnderTest;
            StateUnderTest = stateUnderTest;
            ExpectedBehaviour = expectedBehaviour;
            Cancel = false;
        }

        public string TestClass { get; private set; }
        public IEnumerable<string> MethodsUnderTest { get; private set; }
        public string StateUnderTest { get; private set; }
        public string ExpectedBehaviour { get; private set; }
        public bool Cancel { get; set; }
    }

}