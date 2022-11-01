using System;
using System.CodeDom.Compiler;
using System.Text;

namespace AccessCodeLib.AccUnit
{
    public class CouldNotCompileDynamicTestRowGeneratorException : Exception
    {
        public CouldNotCompileDynamicTestRowGeneratorException(CompilerResults compilerResults)
            : base(GetMessage(compilerResults))
        { }

        private static string GetMessage(CompilerResults compilerResults)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Could not compile the DynamicTestRowGenerator:");
            foreach (CompilerError error in compilerResults.Errors)
            {
                sb.AppendLine(error.ErrorText);
            }
            return sb.ToString();
        }
    }
}