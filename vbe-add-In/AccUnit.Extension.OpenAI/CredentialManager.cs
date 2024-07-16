using System;
using System.Runtime.InteropServices;
using System.Security;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class CredentialManager
    {
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool CredWrite(ref Credential credential, uint reserved);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool CredRead(string target, CredentialType type, uint reserved, out IntPtr credentialPtr);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern void CredFree(IntPtr credential);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct Credential
        {
            public uint Flags;
            public string TargetName;
            public string UserName;
            public SecureString CredentialBlob;
            public uint CredentialBlobSize;
            public CredentialType Type;
            public DateTime LastWritten;
            public uint Persist;
            public uint AttributeCount;
            public IntPtr Attributes;
        }

        private enum CredentialType
        {
            CRED_TYPE_GENERIC = 1
        }

        public void Save(string target, string username, string password)
        {
            var securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }
            securePassword.MakeReadOnly();

            var credential = new Credential
            {
                TargetName = target,
                UserName = username,
                CredentialBlob = securePassword,
                CredentialBlobSize = (uint)securePassword.Length,
                Type = CredentialType.CRED_TYPE_GENERIC
            };

            CredWrite(ref credential, 0);
        }

        public string Retrieve(string target)
        {
            if (CredRead(target, CredentialType.CRED_TYPE_GENERIC, 0, out IntPtr credentialPtr))
            {
                var credential = Marshal.PtrToStructure<Credential>(credentialPtr);
                string password = ConvertToInsecureString(credential.CredentialBlob);
                CredFree(credentialPtr);
                return password;
            }
            return null;
        }

        private string ConvertToInsecureString(SecureString secureString)
        {
            IntPtr ptr = IntPtr.Zero;
            try
            {
                ptr = Marshal.SecureStringToGlobalAllocUnicode(secureString);
                return Marshal.PtrToStringUni(ptr);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(ptr);
            }
        }
    }
}