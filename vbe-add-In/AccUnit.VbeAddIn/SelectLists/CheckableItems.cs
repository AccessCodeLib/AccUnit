using System.Collections.ObjectModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableItems<T> : ObservableCollection<T>
    {
        public new void Add(T item)
        {
            base.Add(item);
            PerformActionOnAddedItem(item);
        }

        protected virtual void PerformActionOnAddedItem(T item)
        {
        }
    }
}
