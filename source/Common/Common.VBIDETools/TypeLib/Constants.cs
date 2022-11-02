using System.Collections.Generic;
using System.Linq;
using TypeLibInformation;

namespace AccessCodeLib.Common.VBIDETools.TypeLib
{
    public class Constants : Dictionary<string, Constant> 
    {
        public Constants()
        {
        }

        public Constants(ConstantInfo tliConstant)
        {
            Add(tliConstant);
        }

        public void Add(ConstantInfo tliConstant)
        {
            ReadMembers(tliConstant);
        }

        private void ReadMembers(ConstantInfo tliConstant)
        {
            foreach (MemberInfo member in tliConstant.Members)
            {
                var c = new Constant(member, tliConstant);
                Add(GetKey(c), c);
            }
        }

        private static string GetKey(Constant constant)
        {
            return string.Format("{0}.{1}.{2}", constant.Name, constant.Parent.Name, constant.Parent.TypeLibInfo.Name);
        }
    }
}