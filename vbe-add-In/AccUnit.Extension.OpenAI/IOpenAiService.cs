using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface IOpenAiService
    {
        Task<string> SendRequest(object[] messages, int maxToken = 0, string model = null);
        string Model { get; set; }
    }
}
