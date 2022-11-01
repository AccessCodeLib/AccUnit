using System.IO;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools
{
    public class CodeModuleInfo
    {

        private readonly string _fileName;

        public CodeModuleInfo()
        {
        }

        public CodeModuleInfo(string name, vbext_ComponentType componentType)
        {
            Name = name;
            ComponentType = componentType;
        }

        public CodeModuleInfo(FileSystemInfo file)
        {
            var pos = file.Name.LastIndexOf(".");
            Name = pos > 0 ? file.Name.Substring(0, pos) : file.Name;
            _fileName = file.FullName;
            ComponentType = GetComponentType(file);
        }

        public static vbext_ComponentType GetComponentType(FileSystemInfo file)
        {
            switch (file.Extension)
            {
                case ".cls":
                    return vbext_ComponentType.vbext_ct_ClassModule;
                case ".bas":
                    return vbext_ComponentType.vbext_ct_StdModule;
                default:
                    return vbext_ComponentType.vbext_ct_Document;
            }
        }

        public string Name { get; set; }
        public vbext_ComponentType ComponentType { get; set; }
        public CodeModuleMemberList Members { get; set; }
        public string FileName { get { return _fileName; } }

        public override string ToString() { return Name; }

    }
}
