namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface ITestCodePromptBuilder
    {
        string BuildPrePrompt(bool useRowTest, string testMethodTemplate);
        string BuildPrompt(string baseProcedureCode, string baseProcedureClassName,
                           string testMethodName, string testMethodParameters);
    }
}