using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.ComponentModel;
using System.Windows.Documents;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestResultViewModel : INotifyPropertyChanged
    {
        public class Result
        {
            public string Message;
            public string Expected;
            public string ButWas;
        }

        private const string ExpectedPrefix = " (Expected: ";
        private const string ButWasPrefix = " but was: ";

        public TestResultViewModel(ITestResult testResult)
        {
            Initialize(testResult);
            CloseCommand = new RelayCommand<Window>(param => CloseWindow(param));
        }

        public string Message { get; private set; }
        public string Expected { get; private set; }
        public string Actual { get; private set; }

        private FlowDocument _highlightedText;

        public FlowDocument HighlightedText
        {
            get { return _highlightedText; }
            set
            {
                _highlightedText = value;
                OnPropertyChanged(nameof(HighlightedText));
            }
        }

        public ICommand CloseCommand { get; }

        private void CloseWindow(object parameter)
        {
            if (parameter is Window window)
            {
                window.Close();
            }
        }

        private void Initialize(ITestResult testResult)
        {
            var result = GetFormattedResult(testResult);
            Message = result.Message ?? string.Empty;
            Expected = result.Expected ?? string.Empty;
            Actual = result.ButWas ?? string.Empty;
            HighlightDifferences(Expected, Actual);
        }

        private Result GetFormattedResult(ITestResult result)
        {
            if (result == null)
                return new Result { Message = "---" };

            var message = result.Message;
            if (string.IsNullOrEmpty(message))
                return new Result { Message = "---" };

            if (message.Substring(0, 2).Equals("  "))
                message = message.Substring(2).Replace("\n  ", "\n");

            message = message.TrimEnd(' ', '\n', '\r');

            return ConvertMessageToResult(message);
        }

        private static Result ConvertMessageToResult(string resultMessage)
        {
            var result = new Result();

            var indexOfExpected = resultMessage.IndexOf(ExpectedPrefix);
            var indexOfButWas = resultMessage.IndexOf(ButWasPrefix);

            if (indexOfExpected == -1 & indexOfButWas == -1)
            {
                result.Message = resultMessage;
                result.Expected = string.Empty;
                result.ButWas = string.Empty;
                return result;
            }

            result.Message = resultMessage.Substring(0, indexOfExpected).Trim(' ', '\r', '\n');
            indexOfExpected += ExpectedPrefix.Length;
            result.Expected = resultMessage.Substring(indexOfExpected, indexOfButWas - indexOfExpected).Trim(' ', '\r', '\n');
            result.ButWas = resultMessage.Substring(indexOfButWas + ButWasPrefix.Length).Trim(' ', '\r', '\n', ')');

            return result;
        }

        private void HighlightDifferences(string expected, string actual)
        {
            FlowDocument document = new FlowDocument();
            Paragraph paragraph = new Paragraph();

            if ( string.IsNullOrEmpty(Expected) || string.IsNullOrEmpty(Actual)
                 || !expected.StartsWith("\"") || !expected.EndsWith("\"") || !actual.StartsWith("\"") || !actual.EndsWith("\""))
            {
                paragraph.Inlines.Add(new Run(actual));
                document.Blocks.Add(paragraph);
                HighlightedText = document; 
                return;
            }

            expected = expected.Substring(1, expected.Length - 2);
            actual = actual.Substring(1, actual.Length - 2);

            int length = Math.Max(expected.Length, actual.Length);

            for (int i = 0; i < length; i++)
            {
                if (i < expected.Length && i < actual.Length && expected[i] == actual[i])
                {
                    paragraph.Inlines.Add(new Run(expected[i].ToString()));
                }
                else
                {
                    if (i < actual.Length)
                    {
                        Run run = new Run(actual[i].ToString());
                        run.Background = Brushes.Yellow;
                        paragraph.Inlines.Add(run);
                    }
                }
            }
 
            paragraph.Inlines.InsertBefore(paragraph.Inlines.FirstInline, new Run("\""));
            paragraph.Inlines.Add(new Run("\""));

            document.Blocks.Add(paragraph);
            HighlightedText = document;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
