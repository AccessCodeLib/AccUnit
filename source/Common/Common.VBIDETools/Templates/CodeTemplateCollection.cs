using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace AccessCodeLib.Common.VBIDETools.Templates
{
    public class CodeTemplateCollection : List<CodeTemplate>
    {

        public CodeTemplateCollection()
        {
        }

        public CodeTemplateCollection(IEnumerable<CodeTemplate> collection)
        {
            AddRange(collection);
        }

        public CodeTemplateCollection(string xmlSerializerString)
        {
            LoadFromString(xmlSerializerString);
        }

        public CodeTemplate this[string name]
        {
            get
            {
                return Find(name);
            }
            set
            {
                var index = FindIndex(x => x.Name == name);
                this[index] = value;
            }
        }

        public CodeTemplate Find(string name)
        {
            return Find(x => x.Name == name);
        }

        public void EnsureModulesExistIn(VBProject vbProject)
        {
            foreach (var codeTemplate in this)
            {
                codeTemplate.EnsureExists(vbProject);
            }
        }

        public void RemoveFromVBProject(VBProject vbProject)
        {
            foreach (var codeTemplate in this)
            {
                codeTemplate.RemoveFromVBProject(vbProject);
            }
        }

        #region XmlSerializer support

        public override string ToString()
        {
            using (new BlockLogger())
            {
                var serializer = new XmlSerializer(typeof(CodeTemplateCollection));
                using (var stringWriter = new StringWriter())
                {
                    serializer.Serialize(stringWriter, this);
                    return stringWriter.ToString();
                }
            }
        }

        public void ExportToFile(string path)
        {
            using (new BlockLogger())
            {
                CheckFolder(path);
                using (var streamWriter = File.CreateText(path))
                {
                    streamWriter.Write(ToString());
                }
            }
        }

        private static void CheckFolder(string filePath)
        {
            using (new BlockLogger())
            {
                var containingFolder = Directory.GetParent(filePath);
                if (!containingFolder.Exists)
                {
                    containingFolder.Create();
                }
            }
        }

        public void LoadFromString(string text)
        {
            using (new BlockLogger())
            {
                // PERF: Instantiating the XmlSerializer takes long
                XmlSerializer serializer;
                using (new BlockLogger("PERF: Instantiating XmlSerializer"))
                {
                    serializer = new XmlSerializer(typeof(CodeTemplateCollection));
                }
                using (new BlockLogger("new StringReader"))
                {
                    using (var stringReader = new StringReader(text))
                    {
                        // PERF: Deserializing takes long as well!
                        object deserialized;
                        using (new BlockLogger("PERF: Deserializing"))
                        {
                            deserialized = serializer.Deserialize(stringReader);
                        }
                        AddRange((CodeTemplateCollection)deserialized);
                    }
                }
            }
        }

        public void LoadFromFile(string path)
        {
            using (new BlockLogger())
            {
                using (var streamReader = File.OpenText(path))
                {
                    LoadFromString(streamReader.ReadToEnd());
                }
            }
        }

        #endregion
    }
}
