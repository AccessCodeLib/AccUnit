using System;
using System.Windows.Forms;
using AccessCodeLib.AccUnit.VbeAddIn.About;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools.Commandbar;
using Microsoft.Office.Core;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class DialogManager : ICommandBarsAdapterClient
    {

        public event EventHandler RefreshTestTemplates;

        private static void ShowAboutDialog()
        {
            var aboutForm = new AboutDialog();
            aboutForm.ShowDialog();
        }

        #region ICommandBarsAdapterClient support

        public void SubscribeToCommandBarAdapter(VbeCommandBarAdapter commandBarAdapter)
        {
            using (new BlockLogger())
            {
                if (!(commandBarAdapter is AccUnitCommandBarAdapter accUnitCommandBarAdapter)) return;

                var popUp = accUnitCommandBarAdapter.AccUnitSubMenu;
                CreateCommandBarItems(commandBarAdapter, popUp);
            }
        }

        private void CreateCommandBarItems(VbeCommandBarAdapter commandBarAdapter, CommandBarPopup popUp)
        {
            var buttonData = new CommandbarButtonData
            {
                Caption = Resources.VbeCommandbars.ToolsUserSettingFormCommandbarButtonCaption,
                Description = string.Empty,
                FaceId = 222,
                BeginGroup = true
            };
            commandBarAdapter.AddCommandBarButton(popUp, buttonData, AccUnitMenuItemsShowUserSettingForm);

            buttonData = new CommandbarButtonData
            {
                Caption = Resources.VbeCommandbars.ToolsAboutCommandbarButtonCaption,
                Description = string.Empty,
                FaceId = 487,
                BeginGroup = true
            };
            commandBarAdapter.AddCommandBarButton(popUp, buttonData, AccUnitMenuItemsShowAboutForm);
        }

        static void AccUnitMenuItemsShowAboutForm(CommandBarButton ctrl, ref bool cancelDefault)
        {
            ShowAboutDialog();
        }

        void AccUnitMenuItemsShowUserSettingForm(CommandBarButton ctrl, ref bool cancelDefault)
        {
            try
            {
                /*
                var dialog = new SettingsWindow(new SettingsViewModel());
                dialog.ShowDialog();
                */
                
                using (var settingDialog = new UserSettingDialog())
                {
                    settingDialog.Settings = UserSettings.Current.Clone();
                    if (settingDialog.ShowDialog() == DialogResult.OK)
                    {
                        UserSettings.Current = settingDialog.Settings as UserSettings;
                        UserSettings.Current?.Save();
                        RefreshTestTemplates?.Invoke(this, null);
                    }
                }
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
            }
        }

        #endregion

    }
}
