using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface IOpenAiService
    {
        void StoreApiKey(string apiKey);
        bool ApiKeyExists();
        Task<string> SendRequest(object[] messages, int maxToken = 0, string model = null);
        string Model { get; set; }
    }
}
