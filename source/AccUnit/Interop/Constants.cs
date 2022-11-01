//using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    //[ComVisible(true)]
    //[Guid("3765BCEC-417A-4153-899F-C9FB25A8310B")]
    public interface IConstantsComInterface
    {
        string DefaultNameSpace { get; }
    }

    //[ComVisible(true)]
    //[Guid("2C7B6C5D-BAC8-4368-AAA6-4763009EAA9D")]
    //[ClassInterface(ClassInterfaceType.None)]
    //[ProgId(ProgIdLibName + ".Constants")]
    public class Constants : IConstantsComInterface
    {
        private const string AccUnitDefaultNameSpace = "AccessCodeLib.AccUnit.interop";

        //[ComVisible(true)]
        public const string ProgIdLibName = "AccUnit";

        //[ComVisible(true)]
        public string DefaultNameSpace { get { return AccUnitDefaultNameSpace; } }
    }

    
}
