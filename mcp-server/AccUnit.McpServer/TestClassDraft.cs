using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>Draft of a test class: the AI fills the test methods from the specification.</summary>
    public class TestClassDraft
    {
        public string TestClassName { get; set; }
        public string ModuleUnderTest { get; set; }
        public string SourceCode { get; set; }

        /// <summary>AccUnit test folder of the database: default folder of accunit_import_test_classes.</summary>
        public string TestFolder { get; set; }
        public List<MethodUnderTest> MethodsUnderTest { get; set; } = new List<MethodUnderTest>();
    }

    public class MethodUnderTest
    {
        public string Name { get; set; }

        /// <summary>Procedure, Property Get, Property Let or Property Set.</summary>
        public string Kind { get; set; }

        /// <summary>Declaration line(s) of the procedure as in the code.</summary>
        public string Declaration { get; set; }

        /// <summary>Comment above the procedure (e.g. NaturalDoc), if any.</summary>
        public string Comment { get; set; }
    }
}
