using AccessCodeLib.AccUnit.Interop;
using NUnit.Framework;
using System;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{
    internal class AssertExceptionTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void RaiseFailedTestAsException()
        {
            NUnit.Framework.Assert.Throws<Interfaces.AssertionException>(() =>
            {
                var assert = new Assertions();
                var Iz = new ConstraintBuilder();
                
                int actual = 1;
                var expected =2;

                assert.That(actual, Iz.EqualTo(expected));
            });
        }

        [Test]
        public void InteropAssertRaiseFailedTestAsException()
        {
            NUnit.Framework.Assert.Throws<Interfaces.AssertionException>(() =>
            {
                var assert = new Interop.Assert();
                var Iz = new Interop.ConstraintBuilder();

                int actual = 1;
                var expected = 2;

                assert.That(actual, Iz.EqualTo(expected));
            });
        }

        [Test]
        public void DontRaiseFailedTestAsExceptionBecauseMatchIsTrue_int_short()
        {
            NUnit.Framework.Assert.DoesNotThrow(() =>
            {
                var assert = new Assertions();
                var Iz = new ConstraintBuilder();
                
                int actual = 1;
                short expected = 1;
                
                assert.That(actual, Iz.EqualTo(expected));
            });
        }

        [Test]
        public void RaiseFailedTestAsException_double_int()
        {
            NUnit.Framework.Assert.Throws<Interfaces.AssertionException>(() =>
            //NUnit.Framework.Assert.DoesNotThrow(() =>
            {
                var assert = new Assertions();
                var Iz = new ConstraintBuilder();
                
                double actual = 1.1;
                int expected = 1;

                assert.That(actual, Iz.EqualTo(expected));
            });
        }

        [Test]
        public void DontRaiseFailedTestAsExceptionBecauseMatchResultCollectorDeactivateThrow()
        {
            NUnit.Framework.Assert.DoesNotThrow(() =>
            {
                var testCollector = new TestCollector();
                var assert = new Assertions
                {
                    MatchResultCollector = testCollector
                };
                var Iz = new ConstraintBuilder();

                var actual = 1;
                var expected = 0;
                
                assert.That(actual, Iz.EqualTo(expected));
            });
        }

        [Test]
        public void RaiseDivideByZeroExceptionAndNotAssertionException()
        {
            try 
            {
                var assert = new Assertions();
                var Iz = new ConstraintBuilder();

                int actual = 1;
                var expected = 1;

                actual = actual / 0;

                assert.That(actual, Iz.EqualTo(expected));
            }
            catch (Exception ex)
            {
                NUnit.Framework.Assert.That(ex, Is.Not.TypeOf<Interfaces.AssertionException>());
                NUnit.Framework.Assert.That(ex, Is.TypeOf<DivideByZeroException>());
            }
        }

        
    }
}
