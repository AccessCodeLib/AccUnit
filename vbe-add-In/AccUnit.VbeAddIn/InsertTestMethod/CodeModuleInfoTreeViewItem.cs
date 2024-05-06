namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CodeModuleInfoTreeViewItem : CheckableTreeViewItemBase<CheckableCodeModuleInfo>
    {
        public CodeModuleInfoTreeViewItem(string fullName, string name, bool isChecked = false)
            : base(fullName, name, isChecked)
        {
        }
    }
}
