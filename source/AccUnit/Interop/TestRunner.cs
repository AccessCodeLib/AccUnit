using AccessCodeLib.AccUnit.Interfaces;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("570D48B4-989D-47CD-852F-F6F8AFE6DD14")]
    public interface ITestRunner : Interfaces.ITestRunner
    {
        /*
         * Run(TestClassInstance, "*")                  ... Alle TestMethoden ausführen
         * Run(TestClassInstance, "MethodenName")       ... Nur einen bestimmten Test ausführen
         * TODO: Run(TestClassInstance, "*Filter*Text*") ... Nur Test, die dem Filterausdruck entsprechen, ausführen
         */
        ITestResult Run([MarshalAs(UnmanagedType.IDispatch)] object TestFixtureInstance, 
                        string TestMethodName = "*", 
                        Interfaces.ITestResultCollector TestResultCollector = null,
                        object filterTags = null);
    }

    [ComVisible(true)]
    [Guid("DBED9DB2-5F34-46A4-87B1-7CB3C4FB94F5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".TestRunner")]
    public class TestRunner : AccUnit.TestRunner.VbaTestRunner, ITestRunner
    {
        public TestRunner(VBProject vbProject = null) : base(vbProject)
        {
        }

        public ITestResult Run([MarshalAs(UnmanagedType.IDispatch)] object TestFixtureInstance, string TestMethodName = "*", 
                                ITestResultCollector TestResultCollector = null,
                                object FilterTags = null)
        {
           
            IEnumerable<ITestItemTag> tags = FilterTags != null ? GetFilterTagEnumerableFromObject(FilterTags) : null;
            return base.Run(TestFixtureInstance, TestMethodName, TestResultCollector, tags);
        }

        public static IEnumerable<ITestItemTag> GetFilterTagEnumerableFromObject(object FilterTags)
        {
            IEnumerable<ITestItemTag> tags = new List<ITestItemTag>();
            if (FilterTags is string)
            {
                if (FilterTags.ToString().Contains(",") || FilterTags.ToString().Contains(";"))
                {
                    // split string into array and add to tags
                    var tagArray = FilterTags.ToString().Split(new char[] { ',', ';' });
                    foreach (var item in tagArray)
                    {
                        ITestItemTag tag = new TestItemTag(item);
                        (tags as List<ITestItemTag>).Add(tag);
                    }
                }
                else
                {
                    ITestItemTag tag = new TestItemTag(FilterTags as string);
                    (tags as List<ITestItemTag>).Add(tag);
                }
            }
            else if (FilterTags is Array)
            {
                foreach (var item in FilterTags as Array)
                {
                    var tag = new TestItemTag(item.ToString());
                    (tags as List<ITestItemTag>).Add(tag);
                }
            }
            else if (FilterTags is IEnumerable<ITestItemTag>)
            {
                (tags as List<ITestItemTag>).AddRange(FilterTags as IEnumerable<ITestItemTag>);
            }

            return tags;
        }
    }
}
