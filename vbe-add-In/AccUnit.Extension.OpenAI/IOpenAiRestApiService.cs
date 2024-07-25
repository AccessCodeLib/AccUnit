using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface IOpenAiRestApiService
    {
        string ApiKey { get; set; }
        Task<string> SendRequest(string jsonRequestBody);
    }
}