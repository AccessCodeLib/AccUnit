using System.ComponentModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public interface ICheckableItem : INotifyPropertyChanged
    {
        bool IsChecked { get; set; }
        string Name { get; set; }
    }
}
