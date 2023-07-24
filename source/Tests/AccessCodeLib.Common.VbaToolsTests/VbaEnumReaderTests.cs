using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.Tools.VBA;
using NUnit.Framework;

namespace AccessCodeLib.Common.VBIDETools.Tests
{
    [TestFixture]
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


            // find int value of vbSunday in enumDictionary without knowing the enum name and return null if not found
            int? testNullable = enumDictionary.Values.SelectMany(x => x).Where(x => x.Key == "vbSundayX").Select(x => (int?)x.Value).FirstOrDefault();
            int test = testNullable ?? -1;
            Assert.That(test, Is.EqualTo(-1));

        }


        [Test]
        public void ReadConstants()
        {
            Dictionary<string, object> dictionary = new Dictionary<string, object>();

            var constants = typeof(VBA.Constants);
            var members = constants.GetMembers();
            foreach (var member in members)
            {
                if (member.MemberType == MemberTypes.Field)
                {
                    object value = constants.GetField(member.Name).GetValue(null);
                    dictionary.Add(member.Name, value);
                }
            }

            foreach (var kvp in dictionary)
            {
                Console.WriteLine($"Name: {kvp.Key}, Wert: {kvp.Value}");
            }

            Assert.That(dictionary["vbNullString"], Is.EqualTo(VBA.Constants.vbNullString));
            Assert.That(dictionary["vbTab"], Is.EqualTo(VBA.Constants.vbTab));
        }

        [Test]
        public void VbaToolsConstantsDictionary_CheckEnums()
        {
            var vbaToolsConstantsDictionary = VbaTools.ConstantsDictionary;

            Assert.That(vbaToolsConstantsDictionary.GetEnumValue("vbFriday"), Is.EqualTo((int)VBA.VbDayOfWeek.vbFriday));
            Assert.That(vbaToolsConstantsDictionary.GetEnumValue("vbfriday"), Is.EqualTo((int)VBA.VbDayOfWeek.vbFriday));

            Assert.That(vbaToolsConstantsDictionary.GetEnumValue("VbDayOfWeek", "vbFriday"), Is.EqualTo((int)VBA.VbDayOfWeek.vbFriday));
            Assert.That(vbaToolsConstantsDictionary.GetEnumValue("VbdayOfweek", "vbfriday"), Is.EqualTo((int)VBA.VbDayOfWeek.vbFriday));
        }

        [Test]
        public void VbaToolsConstantsDictionary_CheckConstants()
        {
            var vbaToolsConstantsDictionary = VbaTools.ConstantsDictionary;

            Assert.That(vbaToolsConstantsDictionary.GetConstantValue("vbNullString"), Is.EqualTo(VBA.Constants.vbNullString));
            Assert.That(vbaToolsConstantsDictionary.GetConstantValue("vbnullstring"), Is.EqualTo(VBA.Constants.vbNullString));

            Assert.That(vbaToolsConstantsDictionary.GetConstantValue("vbTab"), Is.EqualTo(VBA.Constants.vbTab));
            Assert.That(vbaToolsConstantsDictionary.GetConstantValue("vbtab"), Is.EqualTo(VBA.Constants.vbTab));
        }

        [Test]
        [Ignore("only a check, not a test")]
        // check VBA.Constants.vbNullString if that is string pointer 0
        public void CheckPointerOfVbNullString()
        {
            var ptr = Marshal.StringToHGlobalAnsi(VBA.Constants.vbNullString);
            Console.WriteLine($"PTr={ptr}");
            Marshal.FreeHGlobal(ptr);
        }

    }
}
