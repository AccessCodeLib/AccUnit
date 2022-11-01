using System.Text;

namespace AccessCodeLib.Common.Tools.Files
{
    public static class FileTools
    {
        public static string ConvertPathToUNC(string pathToConvert)
        {
            pathToConvert = pathToConvert.Trim();

            var localLetter = GetLocalLetter(pathToConvert); // X:\folder\....
            if (string.IsNullOrEmpty(localLetter))
                return pathToConvert;

            var uncPathFormLocalLetter = GetUncPathFormLocalLetter(localLetter);
            if (string.IsNullOrEmpty(uncPathFormLocalLetter))
                return pathToConvert;

            return uncPathFormLocalLetter + pathToConvert.Substring(localLetter.Length);
        }

        private static string GetUncPathFormLocalLetter(string localLetter)
        {
            var uncStringBuilder = new StringBuilder(255);
            var length = uncStringBuilder.Capacity;
            var result = NativeMethods.WNetGetConnection(localLetter, uncStringBuilder, ref length);
            return result == 0 ? uncStringBuilder.ToString() : null;
        }

        private static string GetLocalLetter(string path)
        {
            var localLetter = path.Substring(0, 2);
            return localLetter.Equals("\\") ? null : localLetter;
        }
    }
}
