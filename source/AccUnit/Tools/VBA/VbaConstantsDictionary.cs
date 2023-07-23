using System;
using System.Reflection;
using VBA;

namespace AccessCodeLib.AccUnit.Tools.VBA
{
    public class VbaConstantsDictionary : AssemblyInfo
    {
        public VbaConstantsDictionary() : base(Assembly.Load("Interop.VBA"))
        {
            FillConstantDictionary();
        }

        private void FillConstantDictionary()
        {
            foreach (Type type in _Assembly.GetTypes())
            {
                if (type.IsClass && type.Name.Contains("Constant"))
                {
                    FillConstantsDictionary(type);
                }
            }
        }
    }
}
