using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableTreeViewItem : CheckableTreeViewItemBase<CheckableItem>
    {
        public CheckableTreeViewItem(string fullName, string name, bool isChecked = false)
            : base(fullName, name, isChecked)
        {
        }
    }

    public class CheckableCodeModulTreeViewItem : CheckableTreeViewItemBase<CheckableCodeModuleMember>
    {
        public CheckableCodeModulTreeViewItem(string fullName, string name, bool isChecked = false)
            : base(fullName, name, isChecked)
        {
        }
    }

    public class CheckableTreeViewItemBase<T> : CheckableItem, ICheckableTreeViewItem<T> 
        where T : CheckableItem
    {
        public CheckableTreeViewItemBase(string fullName, string name, bool isChecked = false)
            : base(fullName, name, isChecked)
        {
            SetChildren();
            Children.CollectionChanged += OnChildrenCollectionChanged;
        }

        public CheckableItems<T> Children { get; set; }

        protected virtual void SetChildren()
        {
            Children = new CheckableItems<T>();
        }

        private void OnChildrenCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                foreach (var item in e.NewItems)
                {
                    if (item is CheckableTreeViewItemBase<T> tvItem)
                    {
                        tvItem.PropertyChanged += OnChildPropertyChanged;
                    }
                }
            }
        }

        private void OnChildPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(IsChecked))
            {
                //if (sender is CheckableTreeViewItemBase<T> tvItem)
                if (sender is ICheckableItem tvItem)
                {
                    if (tvItem.IsChecked)
                    {
                        if (IsChecked == false)
                        {
                            base.SetChecked(true);
                        }
                    }
                    else
                    {
                        foreach (var item in Children)
                        {
                            if (item.IsChecked)
                                return;
                        }
                        base.SetChecked(false);
                    }
                }
            }
        }

        private bool _isExpanded;
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (_isExpanded != value)
                {
                    _isExpanded = value;
                    OnPropertyChanged(nameof(IsExpanded));
                }
            }
        }

        internal override void SetChecked(bool value)
        {
            //SetChecked(value, true);
            base.SetChecked(value);
            ChangeChildrenCheckedState(value);
        }
        /*
        internal virtual void SetChecked(bool value, bool changeChildrenCheckedState)
        {
            base.SetChecked(value);
            if (changeChildrenCheckedState)
            {
                ChangeChildrenCheckedState(value);
            }
        }
        */

        private void ChangeChildrenCheckedState(bool isChecked)
        {
            foreach (var item in Children)
            {
                item.SetChecked(isChecked);
            }
        }

        private ImageSource _imageSource;
        public ImageSource ImageSource
        {
            get { return _imageSource; }
            set
            {
                _imageSource = value;
                OnPropertyChanged(nameof(ImageSource));
            }
        }
    }
}
