// TLB-Export:
/*
 * "$(ProjectDir)..\tools\tlb\tlbExp.exe" $(TargetDir)\$(ProjectName).dll /out:$(TargetDir)\$(ProjectName).tlb
 */

// MarshalAs
/*
 interface IVbaObject
 {
     void Add([MarshalAs(UnmanagedType.IDispatch)] object Object2Add);
 }

 interface IVbaCollections
 {
     [DispId(0)]
     object Item(int Index);

     [DispId(-4)]
     IEnumerable Items();

     [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DISPATCH)]
     object[] ToArray();
 }
*/
