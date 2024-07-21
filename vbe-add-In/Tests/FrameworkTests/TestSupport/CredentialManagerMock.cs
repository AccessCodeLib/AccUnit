using AccessCodeLib.Common.OpenAI;
using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit.Extension.OpenAI.Tests.TestSupport
{
    internal class CredentialManagerMock : ICredentialManager
    {
        private readonly List<CredentialMock> _credentialList = new List<CredentialMock>();

        public string Retrieve(string target)
        {
            var cred = _credentialList.FirstOrDefault<CredentialMock>(m => m.Target == target);
            if (cred == null)
            {
                return string.Empty;
            }
            return cred.Secret;
        }

        public void Save(string target, string username, string secret)
        {
            var cred = _credentialList.FirstOrDefault<CredentialMock>(m => m.Target == target);
            if (cred == null)
            {
                cred = new CredentialMock()
                {
                    Target = target,
                    Username = username,
                    Secret = secret
                };
                _credentialList.Add(cred);
            }
            else
            {
                cred.Username = username;
                cred.Secret = secret;
            }
        }
    }


    public class CredentialMock
    {
        public string Target = string.Empty;
        public string Username = string.Empty;
        public string Secret = string.Empty;
    }
}
