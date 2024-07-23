namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface IOpenAiService
    {
        string SendRequest(object[] messages, int maxToken = 0, string model = null); 
        string Model { get; set; }
    }
}
