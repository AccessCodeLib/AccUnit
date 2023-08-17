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

            IEnumerable<ITestItemTag> tags = FilterTags != null ? GetEnumerableFromFilterObject<ITestItemTag>(FilterTags) : null;
            return base.Run(TestFixtureInstance, TestMethodName, TestResultCollector, tags);
        }

        public static IEnumerable<T> GetEnumerableFromFilterObject<T>(object objectToConvert)
        {
            if (objectToConvert == null)
                return null;    

            IEnumerable<T> tags = new List<T>();
            if (objectToConvert is string)
            {
                if (objectToConvert.ToString().Contains(",") || objectToConvert.ToString().Contains(";"))
                {
                    // split string into array and add to tags
                    var tagArray = objectToConvert.ToString().Split(new char[] { ',', ';' });
                    foreach (var item in tagArray)
                    {
                        var tag = NewItemFromObject<T>(item);
                        (tags as List<T>).Add(tag);
                    }
                }
                else
                {
                    var tag = NewItemFromObject<T>(objectToConvert as string);
                    (tags as List<T>).Add(tag);
                }
            }
            else if (objectToConvert is Array)
            {
                foreach (var item in objectToConvert as Array)
                {
                    var tag = NewItemFromObject<T>(item.ToString());
                    (tags as List<T>).Add(tag);
                }
            }
            else if (objectToConvert is IEnumerable<T>)
            {
                (tags as List<T>).AddRange(objectToConvert as IEnumerable<T>);
            }

            return tags;
        }

        private static T NewItemFromObject<T>(string item)
        {
            if (typeof(T) == typeof(ITestItemTag))
            {
                return (T)(object)new TestItemTag(item);
                //return (T)Activator.CreateInstance(typeof(TestItemTag), item);
            }
                
            else 
                return (T)Activator.CreateInstance(typeof(T), item);
            
        }


    }
}
