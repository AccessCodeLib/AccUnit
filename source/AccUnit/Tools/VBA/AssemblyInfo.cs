using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Collections;

namespace AccessCodeLib.AccUnit.Tools.VBA
{
    public class AssemblyInfo
    {
        protected readonly Assembly _Assembly;
        private readonly Dictionary<string, Dictionary<string, int>> _EnumDictionary = new Dictionary<string, Dictionary<string, int>>();
        private readonly Dictionary<string, Dictionary<string, object>> _ConstantsDictionary = new Dictionary<string, Dictionary<string, object>>();

        public AssemblyInfo(Assembly assembly)
        {
            _Assembly = assembly;
            FillEnumDictionary();
        }

        private void FillEnumDictionary()
        {
            foreach (Type type in _Assembly.GetTypes())
            {
                if (type.IsEnum)
                {
                    FillEnumDictionary(type);
                }
            }
        }

        private void FillEnumDictionary(Type type)
        {
            var enumValues = new Dictionary<string, int>();
            foreach (string enumName in Enum.GetNames(type))
            {
                object enumValue = Enum.Parse(type, enumName);
                enumValues.Add(enumName, (int)enumValue);
            }
            _EnumDictionary.Add(type.Name, enumValues);
        }

        protected void FillConstantsDictionary(Type type)
        {
            var constantsValues = new Dictionary<string, object>();
            var members = type.GetMembers();
            foreach (var member in members)
            {
                if (member.MemberType == MemberTypes.Field)
                {
                    object value = type.GetField(member.Name).GetValue(null);
                    constantsValues.Add(member.Name, value);
                }
            }
            _ConstantsDictionary.Add(type.Name, constantsValues);
        }

        public bool TryGetEnumValue(string enumName, string valueName, out int? value)
        {
            var enums = _EnumDictionary.Select(x => x).Where(x => x.Key.Equals(enumName, StringComparison.InvariantCultureIgnoreCase));

            if (enums.Count() == 0)
            {
                value = null;
                return false;
            }

            var enumDict = enums.First().Value;
            var valuePairs = enumDict.Select(x => x).Where(x => x.Key.Equals(valueName, StringComparison.InvariantCultureIgnoreCase));

            if (valuePairs.Count() == 0)
            {
                value = null;
                return false;
            }

            value = valuePairs.First().Value;
            return true;
        }

        public int? GetEnumValue(string enumName, string valueName)
        {
            if (TryGetEnumValue(enumName, valueName, out int? value))
            {
                return value;
            }
            else
            {
                return null;
            }
        }

        public bool TryGetEnumValue(string enumValueName, out int value)
        {
            var enumValues = GetEnumValue(enumValueName);
            if (enumValues.HasValue)
            {
                value = enumValues.Value;
                return true;
            }
            else
            {
                value = 0;
                return false;
            }
        }

        public int? GetEnumValue(string enumValueName)
        {
            return _EnumDictionary.Values.SelectMany(x => x).Where(x => x.Key.Equals(enumValueName, StringComparison.InvariantCultureIgnoreCase)).Select(x => (int?)x.Value).FirstOrDefault();
        }


        public bool TryGetConstantValue(string constantClass, string constantName, out object value)
        {
            var constants = _ConstantsDictionary.Select(x => x).Where(x => x.Key.Equals(constantClass, StringComparison.InvariantCultureIgnoreCase));

            if (constants.Count() == 0)
            {
                value = null;
                return false;
            }

            var enumDict = constants.First().Value;
            var valuePairs = enumDict.Select(x => x).Where(x => x.Key.Equals(constantName, StringComparison.InvariantCultureIgnoreCase));

            if (valuePairs.Count() == 0)
            {
                value = null;
                return false;
            }

            value = valuePairs.First().Value;
            return true;
        }

        public object GetConstantValue(string constantClass, string constantName)
        {
            if (TryGetConstantValue(constantClass, constantName, out object value))
            {
                return value;
            }
            else
            {
                return null;
            }
        }

        public bool TryGetConstantValue(string constantName, out object value)
        {
            value = GetConstantValue(constantName);
            return value != null;
        }

        public object GetConstantValue(string constantName)
        {
            return _ConstantsDictionary.Values.SelectMany(x => x).Where(x => x.Key.Equals(constantName, StringComparison.InvariantCultureIgnoreCase)).Select(x => x.Value).FirstOrDefault();
        }

    }
}