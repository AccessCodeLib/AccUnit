namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableCodeModuleGroupTreeViewItem : CheckableTreeViewItemBase<CheckableCodeModuleMember>
    {
        public CheckableCodeModuleGroupTreeViewItem(string fullName, string name, bool isChecked = false)
            : base(fullName, name, isChecked)
        {
        }
    }
}
