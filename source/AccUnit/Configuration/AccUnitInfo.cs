using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;

namespace AccessCodeLib.AccUnit.Configuration
{
    public class AccUnitInfo
    {
        public static string FileVersion
        {
            get
            {
                var version = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                return version.FileVersion;
            }
        }

        public static string Copyright
        {
            get
            {
                var version = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                return version.LegalCopyright;
            }
        }
    }
}
