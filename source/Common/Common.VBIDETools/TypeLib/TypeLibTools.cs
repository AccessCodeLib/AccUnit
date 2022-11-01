
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

        public static TLI.Members GetTLIInterfaceMembers(object obj)
        {
            var tliApp = new TLI.TLIApplication();
            var info = tliApp.InterfaceInfoFromObject(obj);
            return info.Members;
        }

    }
}
