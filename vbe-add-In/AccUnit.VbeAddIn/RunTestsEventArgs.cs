namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class RunTestsEventArgs : TestClassListEventArgs
    {
        private readonly bool _breakOnAllErrors;

        public RunTestsEventArgs(TestClassList testclasslist)
            : base(testclasslist)
        {
        }

        public RunTestsEventArgs(TestClassList testclasslist, bool breakOnAllErrors)
            : base(testclasslist)
        {
            _breakOnAllErrors = breakOnAllErrors;
        }

        public bool BreakOnAllErrors
        {
            get { return _breakOnAllErrors; }
        }

    }
}
