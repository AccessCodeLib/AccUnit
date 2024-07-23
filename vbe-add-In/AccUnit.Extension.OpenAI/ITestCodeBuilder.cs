using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface ITestCodeBuilder
    {
        string BuildTestMethodCode();
        string BuildTestMethodCodeAsync();
        ITestCodeBuilder DisableRowTest();
        ITestCodeBuilder ProcedureToTest(string procedureCode, string className = null);
        ITestCodeBuilder TestMethodTemplate(string templateCode);
        ITestCodeBuilder TestMethodName(string testMethodName);
        ITestCodeBuilder TestMethodParameters(string parameterDefinition);
    }
}