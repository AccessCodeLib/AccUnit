using System.Collections.Generic;
using System;
using System.Linq;
using System.Reflection;
using NUnit.Framework;

namespace AccessCodeLib.Common.VbaToolsTests
{
    [TestFixture]
    public class AccessAssemblyTests
    {
        [Test]
        [Ignore("not implented")]
        public void ReadAllEvents()
        {
            Assembly vbaInteropAssembly = Assembly.Load(@"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACC.OLB");

            Dictionary<string, List<string>> eventDictionary = new Dictionary<string, List<string>>();

            foreach (var type in vbaInteropAssembly.GetTypes()) 
            {
                if (type.IsClass)
                {
                    Dictionary<string, int> enumValues = new Dictionary<string, int>();
                    foreach (var eventInfo in type.GetEvents())
                    {
                        eventDictionary[eventInfo.Name].Add(type.Name);
                    }
                }
            }

            
        }
    }
}
