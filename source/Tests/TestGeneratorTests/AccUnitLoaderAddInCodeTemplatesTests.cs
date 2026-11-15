using AccessCodeLib.AccUnit.Configuration;
using NUnit.Framework;
using System.Linq;

namespace TestGeneratorTests
{
    public class AccUnitLoaderAddInCodeTemplatesTests
    {
        private const string SetLoaderFactorySignature = "Public Sub AccUnitFactory_SetLoaderFactory(ByVal Factory As Object)";
        private const string SetLoaderFactoryBody = "Set m_AccUnitLoaderFactory = Factory";

        [TestCase("Microsoft Access")]
        [TestCase("Microsoft Excel")]
        public void FactoryModule_ContainsSetLoaderFactory(string hostAppName)
        {
            var sourceCode = GetFactoryModuleSourceCode(hostAppName);

            Assert.That(sourceCode, Does.Contain(SetLoaderFactorySignature));
            Assert.That(sourceCode, Does.Contain(SetLoaderFactoryBody));
        }

        [TestCase("Microsoft Access")]
        [TestCase("Microsoft Excel")]
        public void FactoryModule_UsesPresetFactoryBeforeLookup(string hostAppName)
        {
            var sourceCode = GetFactoryModuleSourceCode(hostAppName);

            Assert.That(sourceCode, Does.Contain("If m_AccUnitLoaderFactory Is Nothing Then\r\n      Set m_AccUnitLoaderFactory = GetAccUnitLoaderFactory"));
        }

        private static string GetFactoryModuleSourceCode(string hostAppName)
        {
            var templates = new AccUnitLoaderAddInCodeTemplates(false, hostAppName);
            return templates.Single(t => t.Name == "AccUnit_Factory").SourceCode;
        }
    }
}
