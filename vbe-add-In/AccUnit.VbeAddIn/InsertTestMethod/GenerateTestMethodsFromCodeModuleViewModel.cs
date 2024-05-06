using AccessCodeLib.AccUnit.VbeAddIn.TestExplorer;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class GenerateTestMethodsFromCodeModuleViewModel : INotifyPropertyChanged
    {
        public const string TestNamePart_State = "State";
        public const string TestNamePart_Expected = "Expected";

        public delegate void CommitInsertTestMethodsEventHandler(GenerateTestMethodsFromCodeModuleViewModel sender, CommitInsertTestMethodsEventArgs e);
        public event CommitInsertTestMethodsEventHandler InsertTestMethods;

        public event EventHandler Canceled;

        public GenerateTestMethodsFromCodeModuleViewModel(CodeModuleInfo currentCodeModule)
        {
            CurrentCodeModule = currentCodeModule;

            // %ModuleUnderTest%Tests
            TestClassName = AccUnit.Properties.Settings.Default.TestClassNameFormat.Replace("%ModuleUnderTest%", currentCodeModule.Name);

            _stateTestNamePart = new NotifyTestNamePart(TestNamePart_State, Resources.UserControls.InsertTestMethodStateLabelCaption);
            _stateTestNamePart.PropertyChanged += OnNotifyTestNamePartValueChanged;

            _expectedTestNamePart = new NotifyTestNamePart(TestNamePart_Expected, Resources.UserControls.InsertTestMethodExpectedLabelCaption);   
            _expectedTestNamePart.PropertyChanged += OnNotifyTestNamePartValueChanged;

            _methodNameSyntax = new StringValueLabelControlSource("MethodeNameSyntax", "Method name", "<MethodName>_StateUnderTest_ExpectedBehaviour");    

            CancelCommand = new ButtonRelayCommand(Cancel, Resources.UserControls.InsertTestMethodCancelButtonText);
            CommitCommand = new ButtonRelayCommand(Commit, Resources.UserControls.InsertTestMethodCommitButtonText);

            _methods = new CheckableItems<CheckableCodeModulTreeViewItem>
            {
                new CheckableCodeModulTreeViewItem("Methods", "Methods", false),
                new CheckableCodeModulTreeViewItem("Properties", "Properties", false)
            };
            FillMethodsToTest();
        }

        private void FillMethodsToTest()
        {
            var methods = Items.FirstOrDefault(i => i.Name == "Methods");   
            var properties = Items.FirstOrDefault(i => i.Name == "Properties"); 

            foreach (var member in CurrentCodeModule.Members)
            {
                if (member.ProcKind == vbext_ProcKind.vbext_pk_Proc)
                {
                    var item = new CheckableCodeModuleMember(member as CodeModuleMemberWithMarker);

                    methods.Children.Add(item); 
                    if (item.IsChecked)
                    {
                        methods.SetChecked(true);
                        methods.IsExpanded = true;
                    }                  
                }
                else
                {
                    var item = new CheckableCodeModuleMember(member as CodeModuleMemberWithMarker);
                    properties.Children.Add(item);
                    if (item.IsChecked)
                    {
                        properties.SetChecked(true);
                        properties.IsExpanded = true;   
                    }
                }   
            }   
        }

        public CodeModuleInfo CurrentCodeModule { get; set; }   

        public string TestClassName { get; set; }

        private CheckableItems<CheckableCodeModulTreeViewItem> _methods;
        public CheckableItems<CheckableCodeModulTreeViewItem> Items
        {
            get => _methods;
            set
            {
                _methods = value;
                OnPropertyChanged(nameof(Items));
            }
        }

        public IEnumerable<string> MethodNamesToTest 
        { 
            get
            {
                return Items
               .Where(item => item.IsChecked && item is CheckableCodeModulTreeViewItem group)
               .SelectMany(group => group.Children)
               .Where(child => child.IsChecked && child is CheckableCodeModuleMember childMember)
               .Select(childMember => childMember.Name)
               .ToList();
            }
        }

        private void OnNotifyTestNamePartValueChanged(object sender, PropertyChangedEventArgs e)
        {
            FillMethodeNameSyntax();
        }

        private void FillMethodeNameSyntax()
        {
            var statePart = FormatMethodeNamePart(_stateTestNamePart.Value);
            var expectedPart = FormatMethodeNamePart(_expectedTestNamePart.Value);
            _methodNameSyntax.Value = $"<Member>{statePart}{expectedPart}";
            OnPropertyChanged("MethodeNameSyntax");
        }

        private string FormatMethodeNamePart(string partNameValue)
        {
            var part = partNameValue?.Replace(" ", "_");    
            return string.IsNullOrEmpty(part) ? "" : $"_{part}";    
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public IButtonCommand CancelCommand { get; }

        protected void Cancel()
        {
            Canceled?.Invoke(this, EventArgs.Empty);
        }

        public IButtonCommand CommitCommand { get; }

        private readonly INotifyTestNamePart _stateTestNamePart;
        public ITestNamePart StateTestNamePart
        {
            get { return _stateTestNamePart; }
        }

        private readonly INotifyTestNamePart _expectedTestNamePart;
        public ITestNamePart ExpectedTestNamePart
        {
            get { return _expectedTestNamePart; }
        }

        private readonly IStringValueLabelControlSource _methodNameSyntax;
        public IStringValueLabelControlSource MethodNameSyntax
        {
            get { return _methodNameSyntax; }
        }

        protected virtual void Commit()
        {
            InsertTestMethods?.Invoke(this, 
                new CommitInsertTestMethodsEventArgs(TestClassName, MethodNamesToTest, StateTestNamePart.Value, ExpectedTestNamePart.Value));
        }
    }
}
