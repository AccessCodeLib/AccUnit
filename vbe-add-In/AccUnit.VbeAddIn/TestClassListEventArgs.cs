using System;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class TestClassListEventArgs : EventArgs
    {
        protected readonly TestClassList _testClassList;

        public TestClassListEventArgs(TestClassList testclasslist)
        {
            _testClassList = testclasslist;
        }

        public TestClassList TestClassList
        {
            get { return _testClassList; }
        }
    }
}