using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class TestModuleSelectionTests
    {
        private static readonly List<CodeModuleInfo> TestModules = new List<CodeModuleInfo>
        {
            new CodeModuleInfo("modCalcTests", vbext_ComponentType.vbext_ct_ClassModule),
            new CodeModuleInfo("TestHelper", vbext_ComponentType.vbext_ct_StdModule),
            new CodeModuleInfo("Form_frmTestHost", vbext_ComponentType.vbext_ct_Document)
        };

        private static readonly List<string> ComponentNames = new List<string>
        {
            "modCalc", "modCalcTests", "TestHelper", "Form_frmTestHost"
        };

        [Test]
        public void Select_TestClassAndTestRelatedModule_AreSelected()
        {
            var rejected = new List<RejectedTestModule>();

            var selected = TestModuleSelection.Select(new[] { "modcalctests", "TestHelper" }, TestModules, ComponentNames, rejected);

            Assert.That(selected.Select(m => m.Name), Is.EqualTo(new[] { "modCalcTests", "TestHelper" }));
            Assert.That(rejected, Is.Empty);
        }

        [Test]
        public void Select_ModuleThatIsNoTestModule_IsRejected()
        {
            var rejected = new List<RejectedTestModule>();

            var selected = TestModuleSelection.Select(new[] { "modCalc" }, TestModules, ComponentNames, rejected);

            Assert.That(selected, Is.Empty);
            Assert.That(rejected.Single().Reason, Does.Contain("Not a test module"));
        }

        [Test]
        public void Select_UnknownModule_IsRejected()
        {
            var rejected = new List<RejectedTestModule>();

            TestModuleSelection.Select(new[] { "modMissing" }, TestModules, ComponentNames, rejected);

            Assert.That(rejected.Single().Reason, Does.Contain("No module with this name"));
        }

        [Test]
        public void Select_FormModule_IsRejected()
        {
            var rejected = new List<RejectedTestModule>();

            TestModuleSelection.Select(new[] { "Form_frmTestHost" }, TestModules, ComponentNames, rejected);

            Assert.That(rejected.Single().Reason, Does.Contain("not form or report modules"));
        }

        [Test]
        public void Select_SameNameTwice_IsSelectedOnce()
        {
            var selected = TestModuleSelection.Select(new[] { "modCalcTests", "MODCALCTESTS" }, TestModules, ComponentNames,
                                                      new List<RejectedTestModule>());

            Assert.That(selected.Count, Is.EqualTo(1));
        }

        private static readonly object[] MissingNames = { null, new string[0], new[] { " " } };

        [TestCaseSource(nameof(MissingNames))]
        public void Select_NoNames_Throws(string[] names)
        {
            Assert.Throws<DatabaseSessionException>(() =>
                TestModuleSelection.Select(names, TestModules, ComponentNames, new List<RejectedTestModule>()));
        }

        [TestCase(vbext_ComponentType.vbext_ct_ClassModule, ".cls")]
        [TestCase(vbext_ComponentType.vbext_ct_StdModule, ".bas")]
        [TestCase(vbext_ComponentType.vbext_ct_Document, null)]
        public void GetFileExtension_ReturnsExtensionOfVbaExport(vbext_ComponentType componentType, string expected)
        {
            Assert.That(TestModuleSelection.GetFileExtension(componentType), Is.EqualTo(expected));
        }
    }
}
