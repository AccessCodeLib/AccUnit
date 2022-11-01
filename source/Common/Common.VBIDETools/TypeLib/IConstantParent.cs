namespace AccessCodeLib.Common.VBIDETools.TypeLib
{
    public interface IConstantParent
    {
        string Name { get; }
        Constants Constants { get; }
        TypeLibInfo TypeLibInfo { get; }
    }
}