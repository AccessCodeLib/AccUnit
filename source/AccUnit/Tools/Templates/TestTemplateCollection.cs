using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools.Templates;

namespace AccessCodeLib.AccUnit.Tools.Templates
{
    public class TestTemplateCollection : CodeTemplateCollection
    {
        private const string UserTemplatesFileName = "TestTemplates.xml";
        
        public TestTemplateCollection()
        {
            using (new BlockLogger())
            {
                LoadBuiltInTemplates();
                LoadUserTemplates();
            }
        }

        public TestTemplateCollection(IEnumerable<CodeTemplate> collection)
        {
            AddRange(collection);
        }

        public string ExportPath
        {
            get
            {
                using (new BlockLogger())
                {
                    return Path.Combine(GetCheckedTestTemplateFolder(UserSettings.Current.TemplateFolder), UserTemplatesFileName);
                }
            }
        }

        private void LoadUserTemplates()
        {
            using (new BlockLogger())
            {
                var path = ExportPath;
                if (File.Exists(path))
                {
                    LoadFromFile(path);
                }
            }
        }

        private readonly Regex _appFolderReplaceRegex = new Regex("%APPDATA%", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private string GetCheckedTestTemplateFolder(string path)
        {
            using (new BlockLogger())
            {
                path = _appFolderReplaceRegex.Replace(path, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return path;
            }
        }


        private void LoadBuiltInTemplates()
        {
            using (new BlockLogger())
            {
                Logger.Log("===== 1 ========");
                Add(new CodeTemplate(BuiltInTemplateSources.SimpleTestClassName,
                                     BuiltInTemplateSources.SimpleTestClassSource,
                                     BuiltInTemplateSources.SimpleTestClassCaption,
                                     true));
                Logger.Log("===== 2 ========");
                Add(new CodeTemplate(BuiltInTemplateSources.RowTestClassName,
                                     BuiltInTemplateSources.RowTestClassSource,
                                     BuiltInTemplateSources.RowTestClassCaption,
                                     true));
                Logger.Log("===== 3 ========");
            }
        }


        public IEnumerable<CodeTemplate> UserDefinedTemplates
        {
            get
            {
                using (new BlockLogger())
                {
                    return FindAll(x => x.IsBuiltIn == false);
                }
            }
        }

        public void Save()
        {
            var userDefinedTemplates = new CodeTemplateCollection(UserDefinedTemplates);
            userDefinedTemplates.ExportToFile(ExportPath);
        }

    }
}
