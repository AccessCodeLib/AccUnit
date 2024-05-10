using System.ComponentModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public interface ITestNamePart : IStringValueLabelControlSource
    {
    }

    public interface IStringValueLabelControlSource
    {
        string Name { get; }
        string Caption { get; }
        string Value { get; set; }
    }

    public interface INotifyTestNamePart : ITestNamePart, INotifyPropertyChanged
    {
    }
}