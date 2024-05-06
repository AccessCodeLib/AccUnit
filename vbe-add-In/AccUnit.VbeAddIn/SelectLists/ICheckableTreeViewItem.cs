using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public interface ICheckableTreeViewItem : ICheckableTreeViewItem<CheckableItem>
    {
        new CheckableItems<CheckableItem> Children { get; set; }
        new ImageSource ImageSource { get; set; }
        new bool IsExpanded { get; set; }
        new bool IsChecked { get; set; }
        new string Name { get; set; }
    }

    public interface ICheckableTreeViewItem<T> : ICheckableItem
        where T : CheckableItem
    {
        CheckableItems<T> Children { get; set; }
        ImageSource ImageSource { get; set; }
        bool IsExpanded { get; set; }
    }
}