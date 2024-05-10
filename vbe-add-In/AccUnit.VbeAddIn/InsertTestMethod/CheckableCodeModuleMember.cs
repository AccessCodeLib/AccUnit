namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableCodeModuleMember : CheckableItem
    {
        private readonly CodeModuleMemberWithMarker _member;

        public CheckableCodeModuleMember(CodeModuleMemberWithMarker member)
            : base(member.Name, member.Name, member.Marked)
        {
            _member = member;
        }

        internal override void SetChecked(bool value)
        {
            base.SetChecked(value);
            _member.Marked = value;
            OnPropertyChanged(nameof(IsChecked));
        }
    }
}
