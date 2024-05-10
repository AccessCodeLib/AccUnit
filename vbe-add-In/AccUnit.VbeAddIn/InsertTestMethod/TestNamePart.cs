using System.ComponentModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    internal class TestNamePart : StringValueLabelControlSource, ITestNamePart
    {
        public TestNamePart(string name, string caption)
            : base(name, caption)
        {
        }

        public TestNamePart(string name, string caption, string initialValue)
            : base(name, caption, initialValue)
        {
        }
    }

    internal class StringValueLabelControlSource : IStringValueLabelControlSource
    {
        public StringValueLabelControlSource(string name, string caption)
        {
            Name = name;
            Caption = caption;
        }

        public StringValueLabelControlSource(string name, string caption, string initialValue)
        {
            Name = name;
            Caption = caption;
            Value = initialValue;
        }

        public string Name { get; private set; }
        public string Caption { get; private set; }
        public string Value { get; set; }
    }

    internal class NotifyTestNamePart : TestNamePart, INotifyTestNamePart
    {
        public NotifyTestNamePart(string name, string caption)
            : base(name, caption)
        {
        }

        public NotifyTestNamePart(string name, string caption, string initialValue)
            : base(name, caption, initialValue)
        {
        }

        public new string Value
        {
            get { return base.Value; }
            set
            {
                if (base.Value != value)
                {
                    base.Value = value;
                    OnPropertyChanged("Value");
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}