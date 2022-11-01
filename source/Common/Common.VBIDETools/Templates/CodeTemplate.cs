using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing.Design;
using System.Xml.Serialization;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools.Templates
{
    public class CodeTemplate
    {
        public CodeTemplate()
            : this(null, null)
        {
        }

        public CodeTemplate(string name, string sourceCode, string caption = null, bool isBuiltIn = false)
            : this(name, vbext_ComponentType.vbext_ct_ClassModule, sourceCode, caption, isBuiltIn)
        {
        }

        public CodeTemplate(string name, vbext_ComponentType type, string sourceCode, string caption = null, bool isBuiltIn = false)
        {
            using (new BlockLogger())
            {
                Name       = name;
                Type       = type;
                SourceCode = sourceCode;
                IsBuiltIn  = isBuiltIn;

                Caption = caption ?? name;
            }
        }

        [Browsable(false)]
        public bool IsBuiltIn { get; set; }

        [Category("Templates")]
        [XmlAttribute("name")]
        public string Name { get; set; }

        private vbext_ComponentType _type;
        [Browsable(false)]
        [XmlAttribute("type")]
        [DefaultValue(vbext_ComponentType.vbext_ct_ClassModule)]
        public vbext_ComponentType Type { get { return _type; } set { _type = (value != 0 ? value : vbext_ComponentType.vbext_ct_ClassModule); } }

        [Category("Templates")]
        [Editor(typeof(MultilineStringEditor), typeof(UITypeEditor))]
        [XmlAttribute("code")]
        public string SourceCode { get; set; }

        [Category("Templates")]
        [XmlAttribute("caption")]
        public string Caption { get; set; }

        public void EnsureExists(VBProject vbProject)
        {
            using(var codeModuleManager = new CodeModuleContainer(vbProject))
            {
                if (!codeModuleManager.Exists(Name))
                    AddToVBProject(vbProject);
            }
        }

        public CodeModule AddToVBProject(VBProject vbProject, string newName = null)
        {
            using (var codeModuleGenerator = new CodeModuleGenerator(vbProject))
            {
                if (string.IsNullOrEmpty(newName))
                    newName = Name;

                return codeModuleGenerator.Add(Type, newName, SourceCode, true);
            }
        }

        public void RemoveFromVBProject(VBProject vbProject)
        {
            using (var codeModuleManager = new CodeModuleContainer(vbProject))
            {
                codeModuleManager.Remove(Name);
            }
        }
    }
}
