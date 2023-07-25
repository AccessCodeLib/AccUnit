using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Common
{
    internal static class CompareTypeHelper
    {

        public static bool IsIntNumeric(Type T)
        {
            return IntNumericTypes.Contains(Nullable.GetUnderlyingType(T) ?? T);
        }

        private static readonly HashSet<Type> IntNumericTypes = new HashSet<Type>
        {
            typeof(long), typeof(int), typeof(short), typeof(byte), typeof(sbyte),
            typeof(ulong), typeof(uint), typeof(ushort)
        };

        public static bool IsNumeric(Type T)
        {
            return NumericTypes.Contains(Nullable.GetUnderlyingType(T) ?? T);
        }
        private static readonly HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(int),  typeof(double),  typeof(decimal),
            typeof(long), typeof(short),   typeof(sbyte),
            typeof(byte), typeof(ulong),   typeof(ushort),
            typeof(uint), typeof(float)
        };
    }
}
