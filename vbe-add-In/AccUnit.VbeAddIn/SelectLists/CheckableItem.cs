using System.ComponentModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{

    public class CheckableItem : ICheckableItem
    {
        public CheckableItem(string name, bool isChecked = false)
        {
            _name = name;   
            _isChecked = isChecked; 
        }

        private bool _isChecked = false;
        public bool IsChecked
        {
            get { return _isChecked; }
            set
            {
                if (_isChecked != value)
                {
                    SetChecked(value);
                }
            }
        }

        protected virtual void SetChecked(bool value)
        {
            _isChecked = value;
            OnPropertyChanged(nameof(IsChecked));
        }

        private string _name;   
        public string Name {
            get { return _name; }   
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged("Name");
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
