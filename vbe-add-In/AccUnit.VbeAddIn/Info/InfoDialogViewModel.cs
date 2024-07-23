using System.ComponentModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class InfoDialogViewModel : INotifyPropertyChanged
    {
        public InfoDialogViewModel()
        {
        }

        public InfoDialogViewModel(string infoText)
        {
            InfoText = infoText;
        }   

        private string _infoText;
        public string InfoText
        {
            get { return _infoText; }
            set
            {
                _infoText = value;
                OnPropertyChanged(nameof(InfoText));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
