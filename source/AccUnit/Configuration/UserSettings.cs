using AccessCodeLib.AccUnit.Properties;
using AccessCodeLib.AccUnit.Tools.Templates;
using AccessCodeLib.Common.Tools.Logging;
using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing.Design;

namespace AccessCodeLib.AccUnit.Configuration
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
            TestSuiteUserSettings.UnloadCurrent();
            TemplatesUserSettings.UnloadCurrent();
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

        private TestSuiteUserSettings _testSuiteUserSettings;
        private TemplatesUserSettings _toolsUserSettings;

        private UserSettings()
        {
        }

        public UserSettings Clone()
        {
            using (new BlockLogger())
            {
                var clone = new UserSettings
                {
                    _testSuiteUserSettings = TestSuiteUserSettings.Current.Clone(),
                    _toolsUserSettings = TemplatesUserSettings.Current.Clone(),
                    TestClassNameFormat = Settings.Default.TestClassNameFormat,
                };
                return clone;
            }
        }

        #region Load/Save

        private void Load()
        {
            using (new BlockLogger())
            {
                _testSuiteUserSettings = TestSuiteUserSettings.Current.Clone();
                _toolsUserSettings = TemplatesUserSettings.Current.Clone();
                TestClassNameFormat = Settings.Default.TestClassNameFormat;
            }
        }

        public void Save()
        {
            TestSuiteUserSettings.Current = _testSuiteUserSettings;
            TestSuiteUserSettings.Current.Save();
            TemplatesUserSettings.Current = _toolsUserSettings;
            TemplatesUserSettings.Current.Save();

            Settings.Default.TestClassNameFormat = TestClassNameFormat;
        }

        #endregion

        #region Setting Properties

        [Category("Add-in")]
        [DefaultValue("%ModuleUnderTest%Tests")]
        [Description("Naming convention for test classes. Example: %ModuleUnderTest%Tests")]
        public string TestClassNameFormat { get; set; }

        [Category("Add-in")]
        [DefaultValue(false)]
        [Description("Save last state of treeview window on unload and restore window on load (if visible)")]
        // ReSharper disable MemberCanBePrivate.Global
        public bool RestoreVbeWindowsStateOnLoad { get; set; }
        // ReSharper restore MemberCanBePrivate.Global

        #region Tools

        [Category("Import/Export")]
        [DefaultValue(@"%APPFOLDER%\Tests\%APPNAME%")]
        [Description("Import and export folder for test classes\n%APPFOLDER% ... Path to current mdb/accdb\n%APPNAME% ... Filename of mdb/accdb")]
        public string ImportExportFolder
        {
            get { return _testSuiteUserSettings.ImportExportFolder; }
            set { _testSuiteUserSettings.ImportExportFolder = value.TrimEnd('\\', ' ').TrimStart(); }
        }

        [Category("Templates")]
        [DefaultValue(@"%APPDATA%\AccessCodeLib\AccUnit\Templates")]
        [Description("Location of template files")]
        public string TemplateFolder
        {
            get { return _toolsUserSettings.TemplateFolder; }
            set { _toolsUserSettings.TemplateFolder = value.TrimEnd('\\', ' ').TrimStart(); }
        }

        [Browsable(false)]
        public TestTemplateCollection TestTemplates
        {
            get
            {
                using (new BlockLogger())
                {
                    return _toolsUserSettings.TestTemplates;
                }
            }
        }

        [Category("Templates")]
        [Description("Collection of test class templates")]
        public TestTemplateCollection UserDefinedTemplates
        {
            get
            {
                using (new BlockLogger())
                {
                    return _toolsUserSettings.UserDefinedTemplates;
                }
            }
        }

        [Category("Templates")]
        [Description("Template for test methods")]
        [DefaultValue(@"Public Sub {MethodUnderTest}_{StateUnderTest}_{ExpectedBehaviour}({Params})
	' Arrange
	Err.Raise vbObjectError, ""{MethodUnderTest}_{StateUnderTest}_{ExpectedBehaviour}"", ""Test not implemented""
	Const Expected As Variant = ""expected value""
	Dim Actual As Variant
	' Act
	Actual = ""actual value""
	' Assert
	Assert.That Actual, Iz.EqualTo(Expected)
End Sub")]
        [Editor(typeof(MultilineStringEditor), typeof(UITypeEditor))]
        public string TestMethodTemplate
        {
            get { return _toolsUserSettings.TestMethodTemplate; }
            set { _toolsUserSettings.TestMethodTemplate = value; }
        }

        #endregion

        #region TestSuite
        /*
        [Category("Text output")]
        [DefaultValue(60)]
        [Description("")]
        // ReSharper disable MemberCanBePrivate.Global
        public int SeparatorMaxLength
        // ReSharper restore MemberCanBePrivate.Global
        {
            get { return _testSuiteUserSettings.SeparatorMaxLength; }
            set { _testSuiteUserSettings.SeparatorMaxLength = value; }
        }

        [Category("Text output")]
        [DefaultValue('-')]
        [Description("")]
        // ReSharper disable MemberCanBePrivate.Global
        public char SeparatorChar
        // ReSharper restore MemberCanBePrivate.Global
        {
            get { return _testSuiteUserSettings.SeparatorChar; }
            set { _testSuiteUserSettings.SeparatorChar = value; }
        }
        */
        #endregion

        #endregion
    }
}
