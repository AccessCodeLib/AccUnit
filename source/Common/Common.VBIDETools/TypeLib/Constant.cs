/*
 * 
 * "%programfiles(x86)%\Microsoft SDKs\Windows\v7.0A\bin\sgen" /t:AccessCodeLib.Common.VBIDETools.Templates.CodeTemplateCollection /a:"$(TargetPath)" /f
 * 
 */

namespace AccessCodeLib.Common.VBIDETools.TypeLib
{
    public class Constant
    {
        private readonly TLI.MemberInfo _memberInfo;
        private readonly ConstantParent _parent;

        public Constant(TLI.MemberInfo memberInfo, TLI.ConstantInfo parent)
        {
            _memberInfo = memberInfo;
            _parent = new ConstantParent(parent);
        }

        public string Name
        {
            get { return _memberInfo.Name; }
        }

        public object Value
        {
            get { return _memberInfo.Value; }
        }

        public TLI.TliVarType VarType
        {
            get { return _memberInfo.ReturnType.VarType; }
        }

        public IConstantParent Parent
        {
            get { return _parent; }
        }
    }
}