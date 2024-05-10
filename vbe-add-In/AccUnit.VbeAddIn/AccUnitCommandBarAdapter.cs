using AccessCodeLib.AccUnit.VbeAddIn.Properties;
using AccessCodeLib.AccUnit.VbeAddIn.Resources;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools.Commandbar;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class AccUnitCommandBarAdapter : VbeCommandBarAdapter
    {
        private const string AccUnitCommandBarName = "AccUnit";

        public AccUnitCommandBarAdapter(VBE vbe)
            : base(vbe)
        { }

        public CommandBar AccUnitCommandbar { get; private set; }

        public CommandBarPopup AccUnitSubMenu { get; private set; }

        public void Init()
        {
            using (new BlockLogger())
            {
                CreateAccUnitCommandbar();
                SetupAccUnitToolsSubMenu();
            }
        }

        private void CreateAccUnitCommandbar()
        {
            using (new BlockLogger())
            {
                AccUnitCommandbar = VBE.CommandBars.Add(AccUnitCommandBarName, Type.Missing, false, true);
                LoadSettings();
            }
        }

        private void LoadSettings()
        {
            using (new BlockLogger())
            {
                AccUnitCommandbar.Position = (MsoBarPosition)Settings.Default.CommandbarPosition;
                AccUnitCommandbar.RowIndex = Settings.Default.CommandbarRowIndex;
                AccUnitCommandbar.Left = Settings.Default.CommandbarLeft;
                AccUnitCommandbar.Top = Settings.Default.CommandbarTop;
                AccUnitCommandbar.Visible = Settings.Default.CommandbarVisible;
            }
        }

        private void RemoveAccUnitCommandbar()
        {
            try
            {
                SaveSettings(AccUnitCommandbar);
                AccUnitCommandbar.Delete();
                AccUnitCommandbar = null;
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
            }
        }

        private static void SaveSettings(CommandBar accUnitCommandbar)
        {
            Settings.Default.CommandbarPosition = (int)accUnitCommandbar.Position;
            Settings.Default.CommandbarRowIndex = accUnitCommandbar.RowIndex;
            Settings.Default.CommandbarLeft = accUnitCommandbar.Left;
            Settings.Default.CommandbarTop = accUnitCommandbar.Top;
            Settings.Default.CommandbarVisible = accUnitCommandbar.Visible;
            Settings.Default.Save();
        }

        private void SetupAccUnitToolsSubMenu()
        {
            using (new BlockLogger())
            {
                var commandBar = GetToolsCommandBarOrMenuBar();
                AccUnitSubMenu = commandBar.AddPopup();
                // PERF: Reading the settings takes long!
                using (new BlockLogger("PERF: Reading VbeCommandbars.ToolsAccUnitSubMenuCaption"))
                {
                    AccUnitSubMenu.Caption = VbeCommandbars.ToolsAccUnitSubMenuCaption;
                }
                AccUnitSubMenu.BeginGroup = true;
            }
        }

        private CommandBar GetToolsCommandBarOrMenuBar()
        {
            try
            {
                return CommandBarTools;
            }
            catch
            {
                return MenuBar;
            }
        }

        private bool _disposed;

        protected override void Dispose(bool disposing)
        {
            if (_disposed) return;

            try
            {
                if (disposing)
                {
                    // managed resources:
                }

                // unmanaged resources:
                if (AccUnitSubMenu != null)
                {
                    try
                    {
                        AccUnitSubMenu.Delete();
                    }
                    finally
                    {
                        AccUnitSubMenu = null;
                    }
                }


                if (AccUnitCommandbar != null)
                {
                    try
                    {
                        RemoveAccUnitCommandbar();
                    }
                    finally
                    {
                        AccUnitCommandbar = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }
            finally
            {
                base.Dispose(disposing);
            }
            _disposed = true;
        }

    }
}
