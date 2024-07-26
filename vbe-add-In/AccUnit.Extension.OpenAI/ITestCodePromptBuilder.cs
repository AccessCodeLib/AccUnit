namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface ITestCodePromptBuilder
    {
        string BuildPrePrompt(bool useRowTest, string testMethodTemplate);
        string BuildPrompt(string baseProcedureCode, bool isClassMember, string baseProcedureClassName,
                           string testMethodName, string testMethodParameters);
    }
}