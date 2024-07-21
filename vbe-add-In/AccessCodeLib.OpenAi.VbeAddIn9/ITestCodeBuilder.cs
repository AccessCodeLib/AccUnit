namespace AccessCodeLib.OpenAI.VbeAddIn
{
    public interface ITestCodeBuilder
    {
        string BuildTestMethodCode();
        ITestCodeBuilder DisableRowTest();
        ITestCodeBuilder ProcedureToTest(string procedureCode, string className = null);
        ITestCodeBuilder TestMethodTemplate(string templateCode);
        ITestCodeBuilder TestMethodName(string testMethodName);
        ITestCodeBuilder TestMethodParameters(string parameterDefinition);
    }
}