using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Media;
using System.Windows;
using System.ComponentModel;
using System;
using AccessCodeLib.AccUnit.VbeAddIn.TestExplorer;
using System.Windows.Input;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class InsertTestMethodViewModel : INotifyPropertyChanged
    {
        public const string TestNamePart_MethodName = "MethodName";
        public const string TestNamePart_State = "State";
        public const string TestNamePart_Expected = "Expected";

        public delegate void CommitInsertTestMethodEventHandler(InsertTestMethodViewModel sender, TestNamePartsEventArgs e);
        public event CommitInsertTestMethodEventHandler InsertTestMethod;

        public event EventHandler Canceled;

        private readonly ObservableCollection<ITestNamePart> _testNameParts;
        public InsertTestMethodViewModel()
        {
            _testNameParts = new ObservableCollection<ITestNamePart>()
            {
                new TestNamePart(TestNamePart_MethodName, Resources.UserControls.InsertTestMethodMethodNameLabelCaption),
                new TestNamePart(TestNamePart_State, Resources.UserControls.InsertTestMethodStateLabelCaption),
                new TestNamePart(TestNamePart_Expected, Resources.UserControls.InsertTestMethodExpectedLabelCaption)
            };
            MaxCaptionLabelWidth = MeasureCaptionLabelWidth();
            CancelCommand = new RelayCommand(Cancel);
            CommitCommand = new RelayCommand(Commit);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ObservableCollection<ITestNamePart> TestNameParts
        {
            get { return _testNameParts; }
        }

        private int _maxCaptionLabelWidth = 0;
        public int MaxCaptionLabelWidth
        {
            get { 
                if (_maxCaptionLabelWidth == 0)
                {
                    _maxCaptionLabelWidth = MeasureCaptionLabelWidth();
                }
                return  _maxCaptionLabelWidth;
            }
            set
            {
                if (_maxCaptionLabelWidth != value)
                {
                    _maxCaptionLabelWidth = value;
                    OnPropertyChanged("MaxCaptionLabelWidth");
                }
            }
        }

        private int MeasureCaptionLabelWidth()
        {
            double maxCaptionLabelWidth = 50;
            foreach (var testNamePart in _testNameParts)
            {
                var width = MeasureString(testNamePart.Caption);
                if (width > maxCaptionLabelWidth)
                {
                    maxCaptionLabelWidth = width;
                }
            }
            return (int)(Math.Ceiling(maxCaptionLabelWidth) + 20);
        }

        private static double MeasureString(string candidate)
        {
            var typeface = new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
            var formattedText = new FormattedText(candidate, CultureInfo.CurrentUICulture, FlowDirection.LeftToRight, typeface, 12, Brushes.Black, 1.0);
            return formattedText.WidthIncludingTrailingWhitespace;
        }

        public ICommand CancelCommand { get; }

        protected void Cancel()
        {
            Canceled?.Invoke(this, EventArgs.Empty);    
        }

        public ICommand CommitCommand { get; }

        protected virtual void Commit()
        {
            InsertTestMethod?.Invoke(this, new TestNamePartsEventArgs(_testNameParts)); 
        }

    }

    public interface ITestNamePart
    {
        string Name { get; }   
        string Caption { get; } 
        string Value { get; set; }
    }

    internal class TestNamePart : ITestNamePart
    {
        public TestNamePart(string name, string caption)
        {
            Name = name;
            Caption = caption;
        }

        public string Name { get; private set; }
        public string Caption { get; private set; }
        public string Value { get; set; }
    }

    public class TestNamePartsEventArgs : EventArgs
    {
        public TestNamePartsEventArgs()
        {
        }

        public TestNamePartsEventArgs(ICollection<ITestNamePart> items)
        {
            Items = items;
        }

        public ICollection<ITestNamePart> Items { get; set; }
    }

}


