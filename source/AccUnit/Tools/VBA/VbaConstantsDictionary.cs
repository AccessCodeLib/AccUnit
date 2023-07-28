using AccessCodeLib.AccUnit.Properties;
using Microsoft.Vbe.Interop;
using System;
using System.Linq;
using System.Reflection;

namespace AccessCodeLib.AccUnit.Tools.VBA
{
    public class VbaConstantsDictionary : AssemblyInfo
    {
        public VbaConstantsDictionary() : base(GetEmbeddedInterOpAssembly())
        {
            FillConstantDictionary();
        }

        private static Assembly GetEmbeddedInterOpAssembly()
        {
            var assemblyName = new AssemblyName("Interop.VBA");
            var assembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == assemblyName.Name);

            if (assembly is null)
            {
                byte[] interopVbaBytes = Resources.InteropVBA;
                assembly = Assembly.Load(interopVbaBytes);
            }

            return assembly ?? typeof(VBComponent).Assembly;
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
