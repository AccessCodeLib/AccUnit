using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    internal class SelectListBoxViewModel : INotifyPropertyChanged
    {

        public SelectListBoxViewModel()
        {
            Items = new CheckableItemList();
        }

        private CheckableItemList _items;
        public CheckableItemList Items
        {
            get { return _items; }
            set
            {
                if (_items != value)
                {
                    _items = value;
                    OnPropertyChanged("Items");
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
