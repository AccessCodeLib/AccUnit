using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class CredentialManager : ICredentialManager
    {
        [DllImport("Advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredWrite([In] ref Credential userCredential, [In] UInt32 flags);

        [DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredRead(string target, CredentialType type, int reservedFlag, out IntPtr credentialPtr);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern void CredFree(IntPtr credential);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct Credential
        {
            public uint Flags;
            public CredentialType Type;
            public IntPtr TargetName;
            public IntPtr Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public uint CredentialBlobSize;
            public IntPtr CredentialBlob;
            public uint Persist;
            public uint AttributeCount;
            public IntPtr Attributes;
            public IntPtr TargetAlias;
            public IntPtr UserName;
        }

        private enum CredentialPersistence : uint
        {
            Session = 1,
            LocalMachine,
            Enterprise
        }

        public enum CredentialType
        {
            Generic = 1
        }

        public void Save(string target, string username, string secret)
        {
            byte[] byteArray = secret == null ? null : Encoding.Unicode.GetBytes(secret);

            if (byteArray != null && byteArray.Length > 512 * 5)
                throw new ArgumentOutOfRangeException("secret", "The secret message has exceeded 2560 bytes.");

            Credential credential = new Credential
            {
                AttributeCount = 0,
                Attributes = IntPtr.Zero,
                Comment = IntPtr.Zero,
                TargetAlias = IntPtr.Zero,
                Type = CredentialType.Generic,
                Persist = (uint)CredentialPersistence.LocalMachine,
                CredentialBlobSize = (uint)(byteArray?.Length ?? 0),
                TargetName = Marshal.StringToCoTaskMemUni(target),
                CredentialBlob = Marshal.StringToCoTaskMemUni(secret),
                UserName = Marshal.StringToCoTaskMemUni(username ?? Environment.UserName)
            };
            
            bool success = CredWrite(ref credential, 0);
            Marshal.FreeCoTaskMem(credential.TargetName);
            Marshal.FreeCoTaskMem(credential.CredentialBlob);
            Marshal.FreeCoTaskMem(credential.UserName);

            if (!success)
            {
                int lastError = Marshal.GetLastWin32Error();
                throw new Exception(string.Format("CredWrite failed with the error code {0}.", lastError));
            }
        }

        public string Retrieve(string target)
        {
            IntPtr credentialPtr;
            if (!CredRead(target, CredentialType.Generic, 0, out credentialPtr))
            {
                return null;
            }

            var credential = Marshal.PtrToStructure<Credential>(credentialPtr);
            CredFree(credentialPtr);

            string secret = null;
            if (credential.CredentialBlob != IntPtr.Zero)
            {
                secret = Marshal.PtrToStringUni(credential.CredentialBlob, (int)credential.CredentialBlobSize / 2);
            }

            return secret;
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