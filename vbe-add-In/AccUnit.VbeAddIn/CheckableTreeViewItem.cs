using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class CheckableTreeViewItem<T> : CheckableItem
        where T : CheckableItem
    {
        public CheckableTreeViewItem(string fullName, string name, bool isChecked = false)
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
                    if (item is CheckableTreeViewItem<T> tvItem)
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
                if (sender is CheckableTreeViewItem<T> tvItem)
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
            base.SetChecked(value);
            ChangeChildrenCheckedState(value);
        }

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
