using AccessCodeLib.AccUnit.Interfaces;
using System.Collections.ObjectModel;
using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestItem : CheckableTreeViewItem<TestItem>
    {
        public TestItem(string fullName, string name, bool isChecked = false)
            : base(fullName, name, isChecked)
        {
        }

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
