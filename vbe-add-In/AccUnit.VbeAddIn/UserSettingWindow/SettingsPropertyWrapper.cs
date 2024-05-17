using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.VbeAddIn.UserSettingWindow
{
    internal class SettingsPropertyWrapper : INotifyPropertyChanged
    {
        private readonly PropertyInfo _propertyInfo;
        private readonly Properties.Settings _settings;

        public SettingsPropertyWrapper(PropertyInfo propertyInfo, Properties.Settings settings)
        {
            _propertyInfo = propertyInfo;
            _settings = settings;
        }

        public string Name => _propertyInfo.Name;

        public object Value
        {
            get { return _propertyInfo.GetValue(_settings); }
            set
            {
                _propertyInfo.SetValue(_settings, value);
                OnPropertyChanged(nameof(Value));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
