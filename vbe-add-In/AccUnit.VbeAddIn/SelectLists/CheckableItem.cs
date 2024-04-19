using System.ComponentModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableItem : ICheckableItem
    {
        private bool _isChecked;
        public bool IsChecked
        {
            get { return _isChecked; }
            set
            {
                if (_isChecked != value)
                {
                    _isChecked = value;
                    OnPropertyChanged("IsChecked");
                }
            }
        }

        public string Name { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
