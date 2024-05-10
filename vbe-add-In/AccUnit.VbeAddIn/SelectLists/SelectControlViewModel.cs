using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class SelectControlViewModel : INotifyPropertyChanged
    {
        private string _selectAllCheckBoxText = Resources.UserControls.SelectListSelectAllCheckboxCaption;
        private string _commitButtonText = Resources.UserControls.SelectListCommitButtonText;

        public delegate void CommitSelectedItemsEventHandler(SelectControlViewModel sender, CheckableItemsEventArgs e);
        public event CommitSelectedItemsEventHandler ItemsSelected;

        public delegate void RefreshItemListEventHandler(SelectControlViewModel sender, CheckableItemsEventArgs e);
        public event RefreshItemListEventHandler RefreshList;

        public SelectControlViewModel()
        {
            RefreshCommand = new RelayCommand(Refresh);
            CommitCommand = new RelayCommand(Commit);
        }

        public SelectControlViewModel(
                                string title,
                                string selectAllCheckBoxText,
                                string commitButtonText,
                                bool optionalCheckBoxVisibility = false,
                                string optionalCheckBoxText = null)
                : this()
        {
            _selectAllCheckBoxText = selectAllCheckBoxText;
            _commitButtonText = commitButtonText;
            _title = title;
            OptionalCheckboxVisibility = optionalCheckBoxVisibility;
            OptionalCheckBoxText = optionalCheckBoxText;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string _title = "Select Items";
        public string Title
        {
            get { return _title; }
            set
            {
                if (_title != value)
                {
                    _title = value;
                    OnPropertyChanged(nameof(Title));
                }
            }
        }

        private bool _selectAll = false;

        public bool SelectAll
        {
            get { return _selectAll; }
            set
            {
                if (_selectAll != value)
                {
                    _selectAll = value;
                    if (Items != null)
                    {
                        foreach (var item in Items)
                        {
                            item.IsChecked = value;
                        }
                    }
                    OnPropertyChanged(nameof(SelectAll));
                }
            }
        }

        public string SelectAllCheckBoxText
        {
            get => _selectAllCheckBoxText;
            set
            {
                if (_selectAllCheckBoxText != value)
                {
                    _selectAllCheckBoxText = value;
                    OnPropertyChanged(nameof(SelectAllCheckBoxText));
                }
            }
        }

        public string CommitButtonText
        {
            get => _commitButtonText;
            set
            {
                if (_commitButtonText != value)
                {
                    _commitButtonText = value;
                    OnPropertyChanged(nameof(CommitButtonText));
                }
            }
        }

        public ICommand RefreshCommand { get; }
        public ICommand CommitCommand { get; }
        public ImageSource RefreshCommandImageSource
        {
            get
            {
                return UITools.ConvertBitmapToBitmapSource(Resources.Icons.refresh_green);
            }
        }

        protected virtual void Refresh()
        {
            RefreshList?.Invoke(this, new CheckableItemsEventArgs(Items));
        }

        protected virtual void Commit()
        {
            var e = new CheckableItemsEventArgs(SelectedItems);
            var close = true;
            ItemsSelected?.Invoke(this, e);
            if (close)
            {
                //Close();
            }
        }

        public ICollection<ICheckableItem> SelectedItems
        {
            get => Items.Where(i => i.IsChecked).ToList();
        }

        private ObservableCollection<ICheckableItem> _items = new ObservableCollection<ICheckableItem>();
        public ObservableCollection<ICheckableItem> Items
        {
            get { return _items; }
            set
            {
                if (_items != value)
                {
                    _items = value;
                    OnPropertyChanged("Items");
                }
            }
        }



        public string OptionalCheckBoxText { get; private set; }
        public bool OptionalCheckboxVisibility { get; private set; } = false;

        private bool _optionalCheckboxChecked = false;
        public bool OptionalCheckboxChecked
        {
            get { return _optionalCheckboxChecked; }
            set
            {
                _optionalCheckboxChecked = value;
                OnPropertyChanged(nameof(OptionalCheckboxChecked));
            }


        }

    }
}
