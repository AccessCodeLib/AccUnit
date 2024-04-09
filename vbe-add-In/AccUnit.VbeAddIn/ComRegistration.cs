using System;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Win32;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public static class ComRegistration
    {
        public const string ComProgId = "AccUnit.VbeAddIn.Connect";
        private const string HkcuSubKey = @"Software\Microsoft\VBA\VBE\6.0\Addins\" + ComProgId;

        public static void ComRegisterClass(Type t)
        {
            using (new BlockLogger())
            {
                CreateHkcrSubkey(t);
                RegisterVbeAddIn();
            }
        }

        public static void ComUnregisterClass(Type t)
        {
            using (new BlockLogger())
            {
                DeleteHkcuSubkey(HkcuSubKey);
                DeleteHkcrSubkey(GetHkcrSubKey(t));
            }
        }

        private static void CreateHkcrSubkey(Type t)
        {
            using (new BlockLogger())
            {
                var key = Registry.ClassesRoot.CreateSubKey(GetHkcrSubKey(t));
                if (key != null)
                {
                    key.CreateSubKey("Programmable");
                    key.SetValue("", ComProgId);
                    var subkey = key.CreateSubKey(@"InprocServer32\");
                    subkey?.SetValue("", Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
                }
                key = Registry.ClassesRoot.CreateSubKey(ComProgId);
                if (key == null) return;
                key.SetValue("", ComProgId);
                key.Close();
            }
        }

        private static void RegisterVbeAddIn()
        {
            using (new BlockLogger())
            {
                var key = Registry.CurrentUser.CreateSubKey(HkcuSubKey);
                if (key == null) return;
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                key.SetValue("FriendlyName", AddInManager.FriendlyName, RegistryValueKind.String);
                key.SetValue("Description", "AccUnit VBIDE Add-In for Access/VBA", RegistryValueKind.String);
                key.Close();
            }
        }

        private static string GetHkcrSubKey(Type t)
        {
            return "CLSID\\{" + t.GUID.ToString().ToUpper() + "}";
        }

        private static void DeleteHkcuSubkey(string subkey)
        {
            SafeDeleteRegistrySubkey(Registry.CurrentUser, "HKCU", subkey);
        }

        private static void DeleteHkcrSubkey(string subkey)
        {
            SafeDeleteRegistrySubkey(Registry.ClassesRoot, "HKCR", subkey);
        }

        private static void SafeDeleteRegistrySubkey(RegistryKey registryKey, string registryKeyName, string subkey)
        {
            using (new BlockLogger($"Deleting {registryKeyName} sub key {subkey}"))
            {
                try
                {
                    registryKey.DeleteSubKeyTree(subkey);
                }
                catch (Exception exception)
                {
                    Logger.Log(exception);
                }
            }
        }
    }
}