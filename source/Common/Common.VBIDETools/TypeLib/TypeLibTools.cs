using System.Collections.Generic;
using TLI = TypeLibInformation;

namespace AccessCodeLib.Common.VBIDETools.TypeLib
{
    public static class TypeLibTools
    {
        public static string GetTLIInterfaceInfoName(object obj)
        {
            var tliApp = new TLI.TLIApplication();
            var info = tliApp.InterfaceInfoFromObject(obj);
            return info.Name;
        }

        // TODO: Replace with code that not uses TLI-dll
        public static TLI.Members GetTLIInterfaceMembers(object obj)
        {
            var tliApp = new TLI.TLIApplication();
            var info = tliApp.InterfaceInfoFromObject(obj);
            return info.Members;
        }

        // TODO: Replace with code that not uses TLI-dll
        public static IEnumerable<string> GetTLIInterfaceMemberNames(object obj)
        {
            var tliApp = new TLI.TLIApplication();
            var info = tliApp.InterfaceInfoFromObject(obj);
            
            var names = new List<string>();
            foreach (TLI.MemberInfo member in info.Members)
            {
                names.Add(member.Name);
            }
            
            return names;
        }

    }
}
