using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace AccessCodeLib.Common.VBIDETools.Tests
{
    [TestFixture]
    public class AccessAssemblyTests
    {
        [Test]
        public void ReadAllEventsFromAccessLib()
        {
            Assembly vbaInteropAssembly = Assembly.Load(@"Microsoft.Office.Interop.Access");

            Dictionary<string, List<string>> eventDictionary = new Dictionary<string, List<string>>();

            foreach (var type in vbaInteropAssembly.GetTypes())
            {
                if (type.IsInterface || type.IsClass)
                {
                    Dictionary<string, int> enumValues = new Dictionary<string, int>();
                    foreach (var eventInfo in type.GetEvents())
                    {
                        if (!eventDictionary.ContainsKey(eventInfo.Name))
                        {
                            eventDictionary.Add(eventInfo.Name, new List<string>());
                        }
                        eventDictionary[eventInfo.Name].Add(type.Name);
                    }
                }
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\test\events.csv"))
            {
                var maxClasses = eventDictionary.Max(x => x.Value.Count);
                var header = "Event;";
                for (int i = 0; i < maxClasses; i++)
                {
                    header += "Class" + i + ";";
                }
                file.WriteLine(header);
                foreach (var item in eventDictionary)
                {
                    file.WriteLine(item.Key + ";" + string.Join(";", item.Value));
                }
            }
        }
    }
}
