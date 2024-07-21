namespace AccessCodeLib.OpenAI.VbeAddIn
{
    public interface ICredentialManager
    {
        void Save(string target, string username, string secret);
        string Retrieve(string target);
    }
}