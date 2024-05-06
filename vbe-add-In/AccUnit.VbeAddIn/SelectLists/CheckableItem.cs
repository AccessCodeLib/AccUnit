using System.ComponentModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableCodeModuleMember : CheckableItem
    {
        private readonly CodeModuleMemberWithMarker _member;

        public CheckableCodeModuleMember(CodeModuleMemberWithMarker member)
            : base(member.Name, member.Name, member.Marked)
        {
            _member = member;   
        }

        internal override void SetChecked(bool value)
        {
            base.SetChecked(value);
            _member.Marked = value;
            OnPropertyChanged(nameof(IsChecked));
        }


    }

    public class CheckableItem : ICheckableItem
    {
        public CheckableItem(string name, bool isChecked = false)
        {
            _fullName = name;
            _name = name;
            _isChecked = isChecked;
        }

        public CheckableItem(string fullName, string name, bool isChecked = false)
        {
            _fullName = fullName;
            _name = name;
            _isChecked = isChecked;
        }

        private bool _isChecked = false;
        public virtual bool IsChecked
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

        internal virtual void SetChecked(bool value)
        {
            _isChecked = value;
            OnPropertyChanged(nameof(IsChecked));
        }

        private string _fullName;
        public string FullName
        {
            get { return _fullName; }
            set
            {
                if (_fullName != value)
                {
                    _fullName = value;
                    OnPropertyChanged("FullName");
                }
            }
        }

        private string _name;
        public string Name
        {
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
