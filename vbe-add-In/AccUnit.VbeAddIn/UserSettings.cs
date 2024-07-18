using AccessCodeLib.AccUnit.Tools.Templates;
using AccessCodeLib.Common.Tools.Logging;
using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing.Design;

namespace AccessCodeLib.AccUnit.VbeAddIn
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
            AccUnit.UserSettings.UnloadCurrent();
            Tools.UserSettings.UnloadCurrent();
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

        private AccUnit.UserSettings _testSuiteUserSettings;
        private Tools.UserSettings _toolsUserSettings;
        //private AccSpec.Integration.UserSettings _accSpecUserSettings;

        private UserSettings()
        {
        }

        public UserSettings Clone()
        {
            using (new BlockLogger())
            {
                var clone = new UserSettings
                {
                    _testSuiteUserSettings = AccUnit.UserSettings.Current.Clone(),
                    _toolsUserSettings = Tools.UserSettings.Current.Clone(),
                    //_accSpecUserSettings = AccSpec.Integration.UserSettings.Current.Clone(),
                    RestoreVbeWindowsStateOnLoad = RestoreVbeWindowsStateOnLoad,
                    TestClassNameFormat = Properties.Settings.Default.TestClassNameFormat
                };
                return clone;
            }
        }

        #region Load/Save

        private void Load()
        {
            using (new BlockLogger())
            {
                _testSuiteUserSettings = AccUnit.UserSettings.Current.Clone();
                _toolsUserSettings = Tools.UserSettings.Current.Clone();
                //_accSpecUserSettings = AccSpec.Integration.UserSettings.Current.Clone();
                RestoreVbeWindowsStateOnLoad = Properties.Settings.Default.RestoreVbeWindowsStateOnLoad;
                TestClassNameFormat = Properties.Settings.Default.TestClassNameFormat;
                BuildTestMethodsWithChatGPT = Properties.Settings.Default.BuildTestMethodsWithChatGPT;
            }
        }

        public void Save()
        {
            AccUnit.UserSettings.Current = _testSuiteUserSettings;
            AccUnit.UserSettings.Current.Save();
            Tools.UserSettings.Current = _toolsUserSettings;
            Tools.UserSettings.Current.Save();
            /*
             * AccSpec.Integration.UserSettings.Current = _accSpecUserSettings;
            AccSpec.Integration.UserSettings.Current.Save();
            */
            Properties.Settings.Default.RestoreVbeWindowsStateOnLoad = RestoreVbeWindowsStateOnLoad;
            Properties.Settings.Default.TestClassNameFormat = TestClassNameFormat;
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


        [Category("Templates")]
        [Description("Collection of test class templates")]
        public bool BuildTestMethodsWithChatGPT { get; set; }

    //

    #endregion

    #endregion
}
}
