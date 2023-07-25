using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace AccessCodeLib.Common.TestHelpers
{
    public static class EnumerableAssert
    {
        public static void AreEqual<T>(IEnumerable<T> expectedEnumerable, IEnumerable<T> actualEnumerable)
        {
            if (actualEnumerable == null)
            {
                if (expectedEnumerable == null)
                {
                    return;
                }
                Assert.Fail("The actual enumerable is null while the expected is not.");
            }
            if (expectedEnumerable == null)
            {
                Assert.Fail("The actual enumerable is not null while the expected is.");
            }

            var rowCounter = 0;
            var actualEnumerator = actualEnumerable.GetEnumerator();

            foreach (var expectedElement in expectedEnumerable)
            {
                rowCounter++;
                var actualElementIsAvailable = actualEnumerator.MoveNext();
                if (!actualElementIsAvailable)
                {
                    Assert.Fail("The actual enumerable has just {0} elements whereas the expected enumerable has more.",
                                rowCounter - 1);
                }
                Assert.AreEqual(expectedElement, actualEnumerator.Current, "Row number {0} (1-based)", rowCounter);
            }

            if (actualEnumerator.MoveNext())
            {
                Assert.Fail("The actual enumerable has more elements than the expected enumerable which has just {0}.", rowCounter);
            }
        }

        public static void AreOfSpecificTypes<T>(IEnumerable<Type> expectedTypeEnumerable, IEnumerable<T> actualEnumerable)
        {
            if (actualEnumerable == null)
            {
                if (expectedTypeEnumerable == null)
                {
                    return;
                }
                Assert.Fail("The actual enumerable is null while the expected is not.");
            }
            if (expectedTypeEnumerable == null)
            {
                Assert.Fail("The actual enumerable is not null while the expected is.");
            }

            var rowCounter = 0;
            var actualEnumerator = actualEnumerable.GetEnumerator();

            foreach (var expectedType in expectedTypeEnumerable)
            {
                rowCounter++;
                var actualElementIsAvailable = actualEnumerator.MoveNext();
                if (!actualElementIsAvailable)
                {
                    Assert.Fail("The actual enumerable has just {0} elements whereas the expected enumerable has more.",
                                rowCounter - 1);
                }
                Assert.IsInstanceOfType(actualEnumerator.Current, expectedType, "Row number {0} (1-based)", rowCounter);
            }

            if (actualEnumerator.MoveNext())
            {
                Assert.Fail("The actual enumerable has more elements than the expected enumerable which has just {0}.", rowCounter);
            }
        }
    }
}