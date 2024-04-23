using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestClassInfoTestItems : TestItems
    {
        protected override void PerformActionOnAddedItem(TestItem parent)
        {
            var testClassInfoTestItem = (TestClassInfoTestItem)parent;
            var testClassInfo = testClassInfoTestItem.TestClassInfo;
            
            if (testClassInfo.Members == null)
                return; 

            foreach (var member in testClassInfo.Members)
            {
                var testClassMemberInfoTestItem = new TestClassMemberInfoTestItem(member,true);
                parent.Children.Add(testClassMemberInfoTestItem);
            }
        }
    }

    public class TestClassMemberInfoTestItems : TestItems
    {
    }

    public class TestRowTestItems : TestItems
    {
    }

    public class TestItems : ObservableCollection<TestItem> 
    {
        public new void Add(TestItem item)
        {
            base.Add(item);
            PerformActionOnAddedItem(item);   
        }

        protected virtual void PerformActionOnAddedItem(TestItem parent)
        {
        }
    }   

    public class TestClassInfoTestItem : TestItem
    {
        public TestClassInfoTestItem(TestClassInfo testClassInfo, bool isChecked = false)
            : base(testClassInfo.Name, testClassInfo.Name, isChecked)
        {
            TestClassInfo = testClassInfo;
        }

        protected override void SetChildren()
        {
            Children = new TestClassMemberInfoTestItems();
        }

        public TestClassInfo TestClassInfo { get; set; }
    }

    public class TestClassMemberInfoTestItem : TestItem
    {
        public TestClassMemberInfoTestItem(TestClassMemberInfo testClassMemberInfo, bool isChecked = false)
            : base(testClassMemberInfo.FullName, testClassMemberInfo.Name, isChecked)
        {
            TestClassMemberInfo = testClassMemberInfo;
        }

        protected override void SetChildren()
        {
            Children = new TestRowTestItems();
        }

        public TestClassMemberInfo TestClassMemberInfo { get; set; }
    }

    public class TestItem : CheckableItem, INotifyPropertyChanged
    {
        public TestItem(string fullName, string name, bool isChecked = false)
            : base(fullName, name, isChecked)
        {
            SetChildren();
            Children.CollectionChanged += OnChildrenCollectionChanged;
        }

        protected virtual void SetChildren()
        {
            Children = new TestItems();
        }

        private void OnChildrenCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                foreach (var item in e.NewItems)
                {
                    if (item is TestItem testItem)
                    {
                        testItem.PropertyChanged += OnChildPropertyChanged;
                    }
                }
            }
        }

        private void OnChildPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(IsChecked))
            {
                if (sender is TestItem testItem)
                {
                    if (testItem.IsChecked)
                    {
                        if (IsChecked == false)
                        {
                            base.SetChecked(true);
                        }
                    }
                    else
                    {
                        foreach(var item in Children)
                        {
                            if (item.IsChecked)
                                return;
                        }   
                        base.SetChecked(false); 
                    }
                }   
            }
        }

        public TestItems Children { get; set; }

        private ITestResult _testResult;    
        public ITestResult TestResult
        {
            get
            {
                return _testResult;
            }
            set 
            { 
                _testResult = value;
                OnPropertyChanged(nameof(TestResult));
                ImageSource = CalculatedImageSource;

                Result = _testResult == null ? null : (Children.Count == 0 ? _testResult.Message : _testResult.Result);
                OnPropertyChanged(nameof(Result));

                if (_testResult != null && _testResult.IsIgnored && _testResult.Success )
                {
                    if (_testResult is ITestSummary summary)
                    {
                        if (summary.Passed != 0)
                        {
                            return;
                        }   
                    }
                    SetChildsToIgnored();
                }
            }
        }

        private void SetChildsToIgnored()
        {
            foreach (var item in Children)
            {
                item.ImageSource = UITools.ConvertBitmapToBitmapSource(Properties.Resources.noaction_gray); 
            }
        }

        public string Result { get; set; }

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

        private ImageSource CalculatedImageSource
        {
            get
            {
                if (TestResult == null)
                    return null;

                if (TestResult.IsFailure || TestResult.IsError)
                    return UITools.ConvertBitmapToBitmapSource(Properties.Resources.result_failed_16x16);

                if (TestResult.Success)
                {
                    if (TestResult is ITestSummary summary)
                    {
                        if (summary.Passed == 0)
                            return UITools.ConvertBitmapToBitmapSource(Properties.Resources.noaction_gray);
                    }
                    else
                    {
                        if (TestResult.IsIgnored)
                            return UITools.ConvertBitmapToBitmapSource(Properties.Resources.noaction_gray);
                    }

                    return UITools.ConvertBitmapToBitmapSource(Properties.Resources.result_success_16x16);
                }

                if (TestResult.IsPassed)
                    return UITools.ConvertBitmapToBitmapSource(Properties.Resources.result_success_16x16);


                if (TestResult.Executed == false)
                    return UITools.ConvertBitmapToBitmapSource(Properties.Resources.noaction_gray);

                return null;
            }
        }
    }
}
