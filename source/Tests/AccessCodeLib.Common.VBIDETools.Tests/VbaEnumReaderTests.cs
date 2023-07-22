using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using NUnit.Framework;

namespace AccessCodeLib.Common.VBIDETools.Tests
{
    public class VbaEnumReaderTests
    {

        [Test]
        public void ReadEnum()
        {
            Dictionary<string, int> enumDictionary = new Dictionary<string, int>();

            // Verwenden von Reflection, um die Enums aus der VBA-Bibliothek abzurufen
            Type enumType = typeof(VBA.VbDayOfWeek);
            foreach (string enumName in Enum.GetNames(enumType))
            {
                object enumValue = Enum.Parse(enumType, enumName);
                enumDictionary.Add(enumName, (int)enumValue);
            }

            // Das Dictionary ausgeben
            foreach (var kvp in enumDictionary)
            {
                Console.WriteLine($"Name: {kvp.Key}, Wert: {kvp.Value}");
            }

            Assert.That(enumDictionary["vbSunday"], Is.EqualTo((int)VBA.VbDayOfWeek.vbSunday));

        }

        // read all enums from a VBA library
        [Test]
        public void ReadAllEnums()
        {
            Assembly vbaInteropAssembly = Assembly.Load("Interop.VBA");

            Dictionary<string, Dictionary<string, int>> enumDictionary = new Dictionary<string, Dictionary<string, int>>();

            foreach (Type type in vbaInteropAssembly.GetTypes())
            {
                if (type.IsEnum)
                {
                    Dictionary<string, int> enumValues = new Dictionary<string, int>();
                    foreach (string enumName in Enum.GetNames(type))
                    {
                        object enumValue = Enum.Parse(type, enumName);
                        enumValues.Add(enumName, (int)enumValue);
                    }

                    enumDictionary.Add(type.Name, enumValues);
                }
            }

            Assert.That(enumDictionary["VbDayOfWeek"]["vbSunday"], Is.EqualTo((int)VBA.VbDayOfWeek.vbSunday));

            // find int value of vbSunday in enumDictionary without knowing the enum name
            int vbSunday = enumDictionary.Values.SelectMany(x => x).Where(x => x.Key == "vbSunday").Select(x => x.Value).FirstOrDefault();
            Assert.That(vbSunday, Is.EqualTo((int)VBA.VbDayOfWeek.vbSunday));
        }

    }
}
