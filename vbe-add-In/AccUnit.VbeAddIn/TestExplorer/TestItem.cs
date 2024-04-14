using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string FullName { get; set; }
        public ObservableCollection<TestItem> Children { get; set; } = new ObservableCollection<TestItem>();
        public string Result { get; set; }

        private bool _isExpanded;
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (_isExpanded != value)
                {
                    _isExpanded = value;
                    OnPropertyChanged(nameof(IsExpanded));
                }
            }
        }
        public bool IsSelected { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // + Duration, Result, Message ...
    }
}
