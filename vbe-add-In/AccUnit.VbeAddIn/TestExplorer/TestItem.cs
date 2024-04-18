using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestItems : ObservableCollection<TestItem>
    {
    }   

    public class TestItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string FullName { get; set; }
        public TestItems Children { get; set; } = new TestItems();

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
                OnPropertyChanged("ImageSource");
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
        public bool IsSelected { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ImageSource ImageSource
        {
            get
            {
                if (TestResult == null)
                    return null;

                if (TestResult.Success)
                    return UITools.ConvertBitmapToBitmapSource(Properties.Resources.result_success_16x16);

                if (TestResult.IsFailure || TestResult.IsError)
                    return UITools.ConvertBitmapToBitmapSource(Properties.Resources.result_failed_16x16);

                if (TestResult.IsIgnored)
                    return UITools.ConvertBitmapToBitmapSource(Properties.Resources.noaction_gray);

                if (TestResult.Executed == false)
                    return UITools.ConvertBitmapToBitmapSource(Properties.Resources.noaction_gray);

                return null;
            }
        }

        // + Duration, Result, Message ...
    }

}
