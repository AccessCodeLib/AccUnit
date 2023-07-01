using AccessCodeLib.AccUnit.Assertions;
using AccessCodeLib.Common.TestHelpers.AccessRelated;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Office.Interop.Access;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using NUnit.Framework.Internal;
using System;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    public class ArrayAssertTests
    {
        private AccessTestHelper _accessTestHelper;
        private Interop.ITestBuilder _testBuilder;
        private int i;

        [SetUp]
        public void TestBuilderTestsSetup()
        {
            _accessTestHelper = AccessClientTestHelper.NewAccessTestHelper(i++);
            _testBuilder = new Interop.TestBuilder
            {
                HostApplication = _accessTestHelper.Application
            };
        }

        [TearDown]
        public void TestBuilderTestsCleanup()
        {
            if (_testBuilder != null)
            {
                _testBuilder.Dispose();
                _testBuilder = null;
            }

            if (_accessTestHelper != null)
            {
                _accessTestHelper.Dispose();
                _accessTestHelper = null;
            }
        }

        private static Assertions.Assertions NewTestAssert(TestCollector testCollector)
        {
            return new Assertions.Assertions
            {
                MatchResultCollector = testCollector
            };
        }

        [Test]
        public void CheckArray()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function TestMethod() as Variant
   dim X() as variant
   X = Array(1,2,3)
   TestMethod = x     
End Function
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var invocHelper = new InvocationHelper(fixture);
            var actual = invocHelper.InvokeMethod("TestMethod");

            if (actual is object[] actualArray)
            {
                Assert.That(actualArray.Length, Is.EqualTo(3));
                Assert.That(actualArray[0], Is.EqualTo(1));
                Assert.That(actualArray[1], Is.EqualTo(2));
                Assert.That(actualArray[2], Is.EqualTo(3));
            }
        }


        [Test]
        public void IntArrayIsEqual()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function TestMethod() as Long()
   dim X(2) as Long
   x(0) = 1
   x(1) = 2
   x(2) = 3
   TestMethod = x     
End Function
");

            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var invocHelper = new InvocationHelper(fixture);
            var actual = invocHelper.InvokeMethod("TestMethod");

            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            int[] expected = new int[] { 1, 2, 3 };

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }

        [Test]
        public void VariantArrayIsEqual()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function TestMethod() as Variant()
   dim X(2) as Variant
   x(0) = 1
   x(1) = 2
   x(2) = 3
   TestMethod = x     
End Function
");

            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var invocHelper = new InvocationHelper(fixture);
            var actual = invocHelper.InvokeMethod("TestMethod");

            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            //Array expected = new object[] { 1, 2, 3 };
            var expected = invocHelper.InvokeMethod("TestMethod");

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }

        [Test]
        public void ArrayInVariantIsEqualVariantArray()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function ActualArray() as Variant
   ActualArray = Array(1, 2, 3)   
End Function

public Function ExpectedArray() as Variant()
   dim X(2) as Variant
   x(0) = 1
   x(1) = 2
   x(2) = 3
   ExpectedArray = x     
End Function
");

            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var invocHelper = new InvocationHelper(fixture);
            var actual = invocHelper.InvokeMethod("ActualArray");

            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            //Array expected = new object[] { 1, 2, 3 };
            var expected = invocHelper.InvokeMethod("ExpectedArray");

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }

        [Test]
        public void Dim2ArrayInVariantIsEqualVariantArray()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function ActualArray() as Variant
   dim X(2,1) as Variant
   x(0,0) = 1
   x(1,0) = 2
   x(2,0) = 3
   ActualArray = x   
End Function

public Function ExpectedArray() as Variant()
   dim X(2,1) as Variant
   x(0,0) = 1
   x(1,0) = 2
   x(2,0) = 3
   ExpectedArray = x     
End Function
");

            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var invocHelper = new InvocationHelper(fixture);
            var actual = invocHelper.InvokeMethod("ActualArray");

            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            //Array expected = new object[] { 1, 2, 3 };
            var expected = invocHelper.InvokeMethod("ExpectedArray");

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }


    }
}
