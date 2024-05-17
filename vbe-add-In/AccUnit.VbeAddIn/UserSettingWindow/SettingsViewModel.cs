using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AccessCodeLib.AccUnit.VbeAddIn.UserSettingWindow
{
    internal class SettingsViewModel
    {
        public ObservableCollection<SettingsPropertyWrapper> SettingsProperties { get; private set; }
        public ICommand SaveCommand { get; private set; }

        public SettingsViewModel()
        {
            SettingsProperties = new ObservableCollection<SettingsPropertyWrapper>();

            foreach (PropertyInfo property in typeof(Properties.Settings).GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (property.CanRead && property.CanWrite)
                {
                    SettingsProperties.Add(new SettingsPropertyWrapper(property, Properties.Settings.Default));
                }
            }

            SaveCommand = new RelayCommand(SaveSettings);
        }

        public void SaveSettings()
        {
            Properties.Settings.Default.Save();
        }
    }


}
