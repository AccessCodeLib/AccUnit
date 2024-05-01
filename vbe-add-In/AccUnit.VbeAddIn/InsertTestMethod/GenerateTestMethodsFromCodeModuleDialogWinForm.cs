using System;
using System.Linq;
using System.Windows.Forms;
using AccessCodeLib.AccUnit.Tools;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
	public partial class GenerateTestMethodsFromCodeModuleDialogWinForm : Form
	{
	    private static string _testClassNameFormat;
	    private const string ModuleUnderTestPlaceholder = "%ModuleUnderTest%";
        private const string DefaultTestClassNameFormat = "%ModuleUnderTest%Tests";

		private CodeModuleInfo _currentCodeModule;

		public GenerateTestMethodsFromCodeModuleDialogWinForm(string testClassNameFormat)
		{
		    _testClassNameFormat = testClassNameFormat;
		    InitializeComponent();
			tvwMembers.CheckBoxes = true;
			tvwMembers.UseTriStateCheckBoxes = true;

			MethodsNode.SelectedImageKey = MethodsNode.ImageKey;
			PropertiesNode.SelectedImageKey = PropertiesNode.ImageKey;
		}

		public event EventHandler<CommitInsertTestMethodsEventArgs> InsertMethods;

		private TreeNode MethodsNode
		{
			get { return tvwMembers.Nodes["tndMethods"]; }
		}

	    public int MethodsNodesCount
	    {
            get { return MethodsNode.Nodes.Count; }
	    }

		private TreeNode PropertiesNode
		{
			get { return tvwMembers.Nodes["tndProperties"]; }
		}

		public CodeModuleInfo CurrentCodeModule
		{
			get { return _currentCodeModule; }
			set
			{
				_currentCodeModule = value;
				txtModuleUnderTest.Text = _currentCodeModule.Name;
                txtTestclassName.Text = GetTestClassName();
				RefreshMemberList();
			}
		}

	    private string GetTestClassName()
	    {
            return TestClassNameFormat.Replace(ModuleUnderTestPlaceholder, _currentCodeModule.Name);
        }

        private static string TestClassNameFormat
        {
            get
            {
                var result = _testClassNameFormat;
                return string.IsNullOrEmpty(result) ? DefaultTestClassNameFormat : result;
            }
        }

		private void RefreshMemberList()
		{
			ClearMembersInList();
			FillMembers();
		}

		private void ClearMembersInList()
		{
			foreach (var node in tvwMembers.Nodes.Cast<TreeNode>())
			{
				node.Nodes.Clear();
			}
		}

		private void FillMembers()
		{
			FillMembers(MethodsNode, vbext_ProcKind.vbext_pk_Proc);
			FillMembers(PropertiesNode, vbext_ProcKind.vbext_pk_Get);
		}

		private void FillMembers(TreeNode memberGroupNode, vbext_ProcKind procKind)
		{
			var memberNodes = memberGroupNode.Nodes;
			var methods = _currentCodeModule.Members.Where(m => m.ProcKind == procKind);

            tvwMembers.UseTriStateCheckBoxes = true;

		    var markedNodes = 0;

			foreach (var member in methods)
			{
				var newNode = memberNodes.Add(member.Name, member.Name, "Member", "Member");
			    if (!(member is CodeModuleMemberWithMarker))
			    {
			        continue;
			    }

			    newNode.Checked = ((CodeModuleMemberWithMarker)member).Marked;
			    tvwMembers.SetNodeCheckState(newNode, newNode.Checked ? CheckState.Checked : CheckState.Unchecked);
                Logger.Log(string.Format("{0}: Checked = {1}", member.Name, newNode.Checked));

                if (newNode.Checked)
                    markedNodes++;
			}

            if (markedNodes == 0)
                tvwMembers.SetNodeCheckState(memberGroupNode, CheckState.Unchecked);
            else if (markedNodes == memberGroupNode.Nodes.Count)
                tvwMembers.SetNodeCheckState(memberGroupNode, CheckState.Checked);
            else
                tvwMembers.SetNodeCheckState(memberGroupNode, CheckState.Indeterminate);
            
			if (memberNodes.Count > 0)
				memberGroupNode.Expand();
		}

		private void btnInsert_Click(object sender, EventArgs e)
		{
			if (SendInsertMethods()) Close();
		}

		private bool SendInsertMethods()
		{
			if (InsertMethods == null) return false;

			var memberNames = (from member in MethodsNode.Nodes.Cast<TreeNode>() where tvwMembers.GetNodeCheckState(member) == CheckState.Checked select member.Text).ToList();
			memberNames.AddRange(
                (from member in PropertiesNode.Nodes.Cast<TreeNode>() where tvwMembers.GetNodeCheckState(member) == CheckState.Checked select member.Text));

			var e = new CommitInsertTestMethodsEventArgs(txtTestclassName.Text, memberNames,
													     txtStateUnderTest.Text, txtExpectedBehaviour.Text);
			InsertMethods(this, e);
			return !e.Cancel;
		}

		private void txtStateUnderTest_TextChanged(object sender, EventArgs e)
		{
			RefreshMethodNamePreview();
		}

		private void txtExpectedBehaviour_TextChanged(object sender, EventArgs e)
		{
			RefreshMethodNamePreview();
		}

		private void RefreshMethodNamePreview()
		{
			txtMethodNamePreview.Text = string.Format(TestCodeGenerator.TestMethodNameFormat,
													  @"<Member>", txtStateUnderTest.Text, txtExpectedBehaviour.Text);
		}

	}

}
