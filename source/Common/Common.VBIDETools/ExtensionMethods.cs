using System;

namespace AccessCodeLib.Common.VBIDETools
{
    internal static class ExtensionMethods
    {
        public static bool StartsWithLetter(this string value)
        {
            return Char.IsLetter(value[0]);
        }
    }
}