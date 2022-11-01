using System;
using System.ComponentModel;
using AccessCodeLib.AccUnit.Properties;
using AccessCodeLib.Common.Tools.Logging;

namespace AccessCodeLib.AccUnit
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
                if (value == null) throw new ArgumentNullException();
                _current = value;
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
