using AccessCodeLib.AccUnit.Properties;
using AccessCodeLib.AccUnit.Tools.Templates;
using AccessCodeLib.Common.Tools.Logging;
using System;

namespace AccessCodeLib.AccUnit.Tools
{
    public class UserSettings
    {
        #region Static members

        /// <summary>
        /// Unloads the previously loaded instance provided via property Current.
        /// This method is mainly needed to support testing the singleton implementation in property Current.
        /// </summary>
        public static void UnloadCurrent()
        {
            _current = null;
        }

        private static UserSettings _current;
        public static UserSettings Current
        {
            get
            {
                if (_current == null)
                {
                    _current = new UserSettings();
                    _current.Load();
                }
                return _current;
            }
            set
            {
                _current = value ?? throw new ArgumentNullException();
            }
        }

        #endregion

        private UserSettings()
        {
        }

        public UserSettings Clone()
        {
            var clone = new UserSettings
            {
                TemplateFolder = TemplateFolder,
                TestTemplates = TestTemplates,
                UserDefinedTemplates = UserDefinedTemplates,
                TestMethodTemplate = TestMethodTemplate
            };
            return clone;
        }

        #region Load/Save

        private void Load()
        {
            using (new BlockLogger())
            {
                TemplateFolder = Settings.Default.TemplateFolder;
                TestMethodTemplate = GetTestMethodTemplate();
                TestTemplates = new TestTemplateCollection();
                UserDefinedTemplates = new TestTemplateCollection(TestTemplates.UserDefinedTemplates);
            }
        }

        private static string GetTestMethodTemplate()
        {
            var savedTemplate = Settings.Default.TestMethodTemplate;
            return !string.IsNullOrEmpty(savedTemplate) ? savedTemplate : Resources.DefaultTestMethodTemplate;
        }

        public void Save()
        {
            using (new BlockLogger())
            {
                Settings.Default.TestMethodTemplate = TestMethodTemplate;
                Settings.Default.TemplateFolder = TemplateFolder;
                Settings.Default.Save();
                UserDefinedTemplates.Save();
                TestTemplates = new TestTemplateCollection();
            }
        }

        #endregion

        #region Setting Properties

        private string _templateFolder;
        public string TemplateFolder
        {
            get
            {
                return _templateFolder;
            }
            set
            {
                _templateFolder = value.TrimEnd('\\', ' ').TrimStart();
            }
        }

        public TestTemplateCollection TestTemplates { get; private set; }
        public TestTemplateCollection UserDefinedTemplates { get; private set; }
        public string TestMethodTemplate { get; set; }

        #endregion
    }
}
