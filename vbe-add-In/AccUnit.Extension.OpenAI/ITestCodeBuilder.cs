using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface ITestCodeBuilder
    {
        string BuildTestMethodCode();
        Task<string> BuildTestMethodCodeAsync();
        ITestCodeBuilder DisableRowTest();
        ITestCodeBuilder ProcedureToTest(string procedureCode, bool isClassMember, string codeModuleName = null);
        ITestCodeBuilder TestMethodTemplate(string templateCode);
        ITestCodeBuilder TestMethodName(string testMethodName);
        ITestCodeBuilder TestMethodParameters(string parameterDefinition);
    }
}