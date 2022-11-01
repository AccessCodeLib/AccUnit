using AccessCodeLib.Common.TestHelpers.AccessRelated;
using NUnit.Framework;
using Access = Microsoft.Office.Interop.Access;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{
    public class AccessApplicationHelperTests
    {
        [SetUp]
        public void Setup()
        {
            // Is.All
            // Is.EquivalentTo
            // Is.InRange
           
        }

        [Test]
        public void AccessTestHelperApplicationIsNotNullTest()
        {
            var ath = new AccessTestHelper();
            var app = ath.Application as Access.Application;
            Assert.That(app, Is.Not.Null);
        }

        [Test]
        public void CurrentDbIsNotNullTest()
        {
            var ath = new AccessTestHelper();

            Access.Application app = ath.Application;

            var aah = new Common.VBIDETools.AccessApplicationHelper(app);
            var db = aah.CurrentDb;

            ath.Application.Visible = true;

            Assert.That(db, Is.Not.Null);
        }

    }
}