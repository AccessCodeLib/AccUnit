using AccessCodeLib.Common.VBIDETools.Integration;
using NUnit.Framework;

namespace AccessCodeLib.Common.VBIDETools.Tests
{
    [TestFixture]
    public class HostLookupTests
    {
        [TestCase(null, false, ExpectedResult = false, Description = "unknown document: keep active project")]
        [TestCase("", false, ExpectedResult = false, Description = "unknown document: keep active project")]
        [TestCase(@"C:\Test\a.xlsm", true, ExpectedResult = false, Description = "active project matches: keep it")]
        [TestCase(@"C:\Test\a.xlsm", false, ExpectedResult = true, Description = "active project does not match: look up")]
        public bool IsProjectLookupRequired_ReturnsExpected(string fullName, bool activeProjectMatches)
        {
            return OfficeApplicationHelper.IsProjectLookupRequired(fullName, _ => activeProjectMatches);
        }

        [Test]
        public void IsProjectLookupRequired_EmptyName_DoesNotEvaluateMatch()
        {
            var evaluated = false;

            OfficeApplicationHelper.IsProjectLookupRequired(string.Empty, _ => evaluated = true);

            Assert.That(evaluated, Is.False);
        }

        [Test]
        public void GetOfficeComObjectByProgIDAndVBE_ProgIdNotRegistered_ReturnsNullInsteadOfThrowing()
        {
            // not "Access.Application" => the VBE is not touched; Marshal.GetActiveObject throws for an unknown ProgID
            object result = null;

            Assert.DoesNotThrow(() => result = ComTools.GetOfficeComObjectByProgIDAndVBE("AccUnit.NotRegistered.Application", null));

            Assert.That(result, Is.Null);
        }
    }
}
