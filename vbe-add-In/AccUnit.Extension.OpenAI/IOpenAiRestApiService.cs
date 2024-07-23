namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface IOpenAiRestApiService
    {
        string ApiKey { get; set; }

        string SendRequest(string jsonRequestBody);
    }
}