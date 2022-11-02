namespace AccessCodeLib.Common.VBIDETools.TypeLib
{
    internal class ConstantParent : IConstantParent
    {
        private readonly TLI.ConstantInfo _constantInfo;

        public ConstantParent(TLI.ConstantInfo constantInfo)
        {
            _constantInfo = constantInfo;
        }

        public string Name { get { return _constantInfo.Name; } }
        public Constants Constants { get { return new Constants(_constantInfo); } }
        public TypeLibInfo TypeLibInfo { get { return new TypeLibInfo(_constantInfo.Parent); } }
    }



}