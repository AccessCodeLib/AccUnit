using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public interface ICheckableTreeViewItem<T> where T : CheckableItem
    {
        CheckableItems<T> Children { get; set; }
        ImageSource ImageSource { get; set; }
        bool IsExpanded { get; set; }
    }
}