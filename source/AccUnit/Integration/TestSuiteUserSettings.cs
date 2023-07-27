using AccessCodeLib.AccUnit.Properties;
using AccessCodeLib.Common.Tools.Logging;
using System;
using System.ComponentModel;

namespace AccessCodeLib.AccUnit
{
    public class TestSuiteUserSettings
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

        private static TestSuiteUserSettings _current;
        public static TestSuiteUserSettings Current
        {
            get
            {
                if (_current == null)
                {
                    _current = new TestSuiteUserSettings();
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

        private TestSuiteUserSettings()
        {
        }

        public TestSuiteUserSettings Clone()
        {
            var clone = new TestSuiteUserSettings
            {
                ImportExportFolder = ImportExportFolder,
                SeparatorChar = SeparatorChar,
                SeparatorMaxLength = SeparatorMaxLength
            };
            return clone;
        }

        #region Load/Save

        private void Load()
        {
            using (new BlockLogger())
            {
                ImportExportFolder = Settings.Default.ImportExportFolder;
                SeparatorMaxLength = Settings.Default.SeparatorMaxLength;
                SeparatorChar = Settings.Default.SeparatorChar;
            }
        }

        public void Save()
        {
            Settings.Default.ImportExportFolder = ImportExportFolder;
            Settings.Default.SeparatorMaxLength = SeparatorMaxLength;
            Settings.Default.SeparatorChar = SeparatorChar;
            Settings.Default.Save();
        }

        #endregion

        #region Setting Properties

        private string _importExportFolder;
        public string ImportExportFolder
        {
            get
            {
                return _importExportFolder;
            }
            set
            {
                _importExportFolder = value.TrimEnd('\\', ' ').TrimStart();
            }
        }

        [Category("Text output")]
        [DefaultValue(60)]
        [Description("")]
        // ReSharper disable MemberCanBePrivate.Global
        public int SeparatorMaxLength { get; set; }
        // ReSharper restore MemberCanBePrivate.Global

        [Category("Text output")]
        [DefaultValue('-')]
        [Description("")]
        // ReSharper disable MemberCanBePrivate.Global
        public char SeparatorChar { get; set; }
        // ReSharper restore MemberCanBePrivate.Global


        #endregion
    }
}
