using System;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class GetTestClassInfoEventArgs : EventArgs
    {
        public GetTestClassInfoEventArgs(string className)
        {
            ClassName = className;
        }

        public string ClassName { get; }   
        public TestClassInfo TestClassInfo { get; set; }
    }
}