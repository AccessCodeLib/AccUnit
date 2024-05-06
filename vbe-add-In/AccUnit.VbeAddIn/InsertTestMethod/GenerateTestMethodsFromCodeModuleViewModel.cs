using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class GenerateTestMethodsFromCodeModuleViewModel : INotifyPropertyChanged
    {
        public const string TestNamePart_State = "State";
        public const string TestNamePart_Expected = "Expected";

        private const string ModuleUnderTestPlaceholder = "%ModuleUnderTest%";

        public delegate void CommitInsertTestMethodsEventHandler(GenerateTestMethodsFromCodeModuleViewModel sender, CommitInsertTestMethodsEventArgs e);
        public event CommitInsertTestMethodsEventHandler InsertTestMethods;

        public event EventHandler Canceled;

        public GenerateTestMethodsFromCodeModuleViewModel(CodeModuleInfo currentCodeModule)
        {
            CurrentCodeModule = currentCodeModule;

            _codeModuleToTestInfo = new StringValueLabelControlSource(
                                            "CodeModuleToTestInfo", 
                                             Resources.UserControls.InsertTestMethodsCodeModuleToTestLabelCaption
                                             , currentCodeModule.Name);

            // %ModuleUnderTest%Tests
            var testClassName = AccUnit.Properties.Settings.Default.TestClassNameFormat.Replace(ModuleUnderTestPlaceholder, currentCodeModule.Name);
            _testClassName = new StringValueLabelControlSource("TestClassName", "Test class", testClassName);

            _stateTestNamePart = new NotifyTestNamePart(TestNamePart_State, Resources.UserControls.InsertTestMethodStateLabelCaption);
            _stateTestNamePart.PropertyChanged += OnNotifyTestNamePartValueChanged;

            _expectedTestNamePart = new NotifyTestNamePart(TestNamePart_Expected, Resources.UserControls.InsertTestMethodExpectedLabelCaption);   
            _expectedTestNamePart.PropertyChanged += OnNotifyTestNamePartValueChanged;

            _methodNameSyntax = new StringValueLabelControlSource("MethodeNameSyntax", "Test name", "<MethodName>_{State}_{Expected}");    

            CancelCommand = new ButtonRelayCommand(Cancel, Resources.UserControls.InsertTestMethodCancelButtonText);
            CommitCommand = new ButtonRelayCommand(Commit, Resources.UserControls.InsertTestMethodCommitButtonText);

            _memberGroups = new CheckableItems<CheckableCodeModuleGroupTreeViewItem>();
            AppendMemberGroups();
        }

        public string SelectedModuleInstruction => Resources.UserControls.InsertTestMethodsSelectedModuleInstruction;
        public string SelectMemberCaption => Resources.UserControls.InsertTestMethodsSelectMemberCaption;  

        public CodeModuleInfo CurrentCodeModule { get; set; }

        private readonly IStringValueLabelControlSource _codeModuleToTestInfo;
        public IStringValueLabelControlSource CodeModuleToTestInfo
        {
            get { return _codeModuleToTestInfo; }
        }

        //public string TestClassName { get; set; }
        private readonly IStringValueLabelControlSource _testClassName;
        public IStringValueLabelControlSource TestClassName
        {
            get { return _testClassName; }
        }

        private CheckableItems<CheckableCodeModuleGroupTreeViewItem> _memberGroups;
        public CheckableItems<CheckableCodeModuleGroupTreeViewItem> Items
        {
            get => _memberGroups;
            set
            {
                _memberGroups = value;
                OnPropertyChanged(nameof(Items));
            }
        }

        private void AppendMemberGroups()
        {
            var methods = CurrentCodeModule.Members.Where(m => m.ProcKind == vbext_ProcKind.vbext_pk_Proc);
            if (methods.Any())
            {
                methods = methods.OrderBy(p => p.Name);
                var methodsGroup = new CheckableCodeModuleGroupTreeViewItem("Methods", "Methods", false);                
                _memberGroups.Add(methodsGroup);
                FillMemberGroup(methodsGroup, methods);
            }

            var properties = CurrentCodeModule.Members.Where(m => m.ProcKind == vbext_ProcKind.vbext_pk_Get);
            properties = AppendProperties(properties, vbext_ProcKind.vbext_pk_Let);
            properties = AppendProperties(properties, vbext_ProcKind.vbext_pk_Set);
            if (properties.Any())
            {
                properties = properties.OrderBy(p => p.Name);
                var propertiesGroup = new CheckableCodeModuleGroupTreeViewItem("Properties", "Properties", false);
                _memberGroups.Add(propertiesGroup);
                FillMemberGroup(propertiesGroup, properties);
            }   
        }

        private IEnumerable<CodeModuleMember> AppendProperties(IEnumerable<CodeModuleMember> members, vbext_ProcKind procKind)
        {
            var newMembers = CurrentCodeModule.Members.Where(m => m.ProcKind == procKind);
            if (!newMembers.Any())
            {
                return members;
            }

            var membersToAppend = newMembers.Where(p => !members.Any(pr => pr.Name == p.Name));
            if (!membersToAppend.Any())
            {
                return members;
            }
            return members.Concat(membersToAppend);
        }

        private void FillMemberGroup(CheckableCodeModuleGroupTreeViewItem group, IEnumerable<CodeModuleMember> members)
        {
            if (members == null || !members.Any())
            {
                return;
            }

            foreach (var member in members)
            {
                AddMemberToGroup(group, member);
            }

        }

        private void AddMemberToGroup(CheckableCodeModuleGroupTreeViewItem group, CodeModuleMember member)
        {
            var item = new CheckableCodeModuleMember(member as CodeModuleMemberWithMarker);
            group.Children.Add(item);
            if (item.IsChecked && !group.IsChecked)
            {
                group.SetChecked(true, false);
                group.IsExpanded = true;
            }
        }

        public IEnumerable<string> MethodNamesToTest 
        { 
            get
            {
                return Items
               .Where(item => item.IsChecked && item is CheckableCodeModuleGroupTreeViewItem group)
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
            OnPropertyChanged(nameof(MethodNameSyntax));
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
                new CommitInsertTestMethodsEventArgs(TestClassName.Value, MethodNamesToTest, StateTestNamePart.Value, ExpectedTestNamePart.Value));
        }
    }
}
