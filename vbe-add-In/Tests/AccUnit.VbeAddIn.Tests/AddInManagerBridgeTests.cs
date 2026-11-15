using NUnit.Framework;
using System;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace AccessCodeLib.AccUnit.VbeAddIn.Tests
{
    /// <summary>
    /// VBA uses late binding on the object returned by AddInManagerBridge.TestSuite.
    /// These tests call the bridge through its IDispatch, like VBA does, and check that
    /// the members used by the AccUnit_Factory module can be resolved on the returned object.
    /// </summary>
    public class AddInManagerBridgeTests
    {
        private static readonly string[] TestSuiteMemberNames =
        {
            "Reset", "Add", "AddByClassName", "AddFromVBProject",
            "SelectTests", "Filter", "Run", "Summary", "Name", "AppendTestResultReporter"
        };

        [TestCase(typeof(Interop.AccessTestSuite))]
        [TestCase(typeof(Interop.VBATestSuite))]
        public void TestSuite_CalledViaIDispatch_ReturnedObjectSupportsTestSuiteMembers(Type testSuiteType)
        {
            var suite = CreateTestSuiteWithoutHostApplication(testSuiteType);
            var bridge = new AddInManagerBridge();
            bridge.TestSuiteRequest += (out Interfaces.IVBATestSuite testSuite) => testSuite = suite;

            var bridgeDispatch = Marshal.GetIDispatchForObject(bridge);
            try
            {
                var suiteDispatch = DispatchHelper.InvokeReturningDispatch(bridgeDispatch, "TestSuite", (int)TestReportOutput.DebugPrint);
                try
                {
                    foreach (var memberName in TestSuiteMemberNames)
                    {
                        Assert.That(DispatchHelper.TryGetDispId(suiteDispatch, memberName), Is.True,
                                    $"Member '{memberName}' is not available via IDispatch");
                    }
                }
                finally
                {
                    Marshal.Release(suiteDispatch);
                }
            }
            finally
            {
                Marshal.Release(bridgeDispatch);
            }
        }

        private static readonly string[] FactoryMemberNamesUsedByTemplates =
        {
            "Assert", "ConstraintBuilder", "TestSuite", "AccUnitFactory", "CodeCoverageTracker", "ErrorTrappingObserver"
        };

        [Test]
        public void BridgeViaIDispatch_ProvidesAllFactoryMembersUsedByTemplates()
        {
            var bridgeDispatch = Marshal.GetIDispatchForObject(new AddInManagerBridge());
            try
            {
                foreach (var memberName in FactoryMemberNamesUsedByTemplates)
                {
                    Assert.That(DispatchHelper.TryGetDispId(bridgeDispatch, memberName), Is.True,
                                $"Member '{memberName}' is not available via IDispatch");
                }
            }
            finally
            {
                Marshal.Release(bridgeDispatch);
            }
        }

        [Test]
        public void AccUnitFactory_CalledViaIDispatch_ReturnsDispatchWithFactoryMembers()
        {
            var bridge = new AddInManagerBridge();
            bridge.AccUnitFactoryRequest += (out Interop.IAccUnitFactory accUnitFactory) => accUnitFactory = new Interop.AccUnitFactory();

            AssertBridgeMemberReturnsDispatchWithMembers(bridge, "AccUnitFactory",
                "Assert", "ConstraintBuilder", "CodeCoverageTracker", "AccessErrorTrappingObserver");
        }

        [Test]
        public void CodeCoverageTracker_CalledViaIDispatch_ReturnsDispatchWithTrackerMembers()
        {
            var bridge = new AddInManagerBridge();
            bridge.CodeCoverageTrackerRequest += (out Interop.ICodeCoverageTracker tracker) => tracker = new Interop.CodeCoverageTracker();

            AssertBridgeMemberReturnsDispatchWithMembers(bridge, "CodeCoverageTracker",
                "Add", "Track", "GetReport", "Clear", "Dispose");
        }

        [Test]
        public void ErrorTrappingObserver_CalledViaIDispatch_ReturnsDispatchWithObserverMembers()
        {
            // A real observer reads the Access option "Error Trapping"; only the COM type information is needed here.
            var observer = FormatterServices.GetUninitializedObject(typeof(Configuration.AccessErrorTrappingObserver));
            GC.SuppressFinalize(observer);
            var bridge = new AddInManagerBridge();
            bridge.ErrorTrappingObserverRequest += (out Configuration.IErrorTrappingObserver errorTrappingObserver) =>
                errorTrappingObserver = (Configuration.IErrorTrappingObserver)observer;

            AssertBridgeMemberReturnsDispatchWithMembers(bridge, "ErrorTrappingObserver",
                "SetErrorTrapping", "Dispose");
        }

        private static void AssertBridgeMemberReturnsDispatchWithMembers(AddInManagerBridge bridge, string bridgeMemberName, params string[] expectedMemberNames)
        {
            var bridgeDispatch = Marshal.GetIDispatchForObject(bridge);
            try
            {
                var resultDispatch = DispatchHelper.InvokeReturningDispatch(bridgeDispatch, bridgeMemberName);
                try
                {
                    foreach (var memberName in expectedMemberNames)
                    {
                        Assert.That(DispatchHelper.TryGetDispId(resultDispatch, memberName), Is.True,
                                    $"Member '{memberName}' of '{bridgeMemberName}' is not available via IDispatch");
                    }
                }
                finally
                {
                    Marshal.Release(resultDispatch);
                }
            }
            finally
            {
                Marshal.Release(bridgeDispatch);
            }
        }

        private static Interfaces.IVBATestSuite CreateTestSuiteWithoutHostApplication(Type testSuiteType)
        {
            // Only the COM type information is needed; a real suite would require an Office host application.
            var suite = FormatterServices.GetUninitializedObject(testSuiteType);
            GC.SuppressFinalize(suite);
            return (Interfaces.IVBATestSuite)suite;
        }

        private static class DispatchHelper
        {
            private const int GetIDsOfNamesVtableSlot = 5;
            private const int InvokeVtableSlot = 6;
            private const short DispatchMethodOrPropertyGet = 1 | 2;
            private const short VariantTypeDispatch = 9;
            private static readonly int VariantSize = IntPtr.Size == 8 ? 24 : 16; // sizeof(VARIANT)
            private const int LocaleUserDefault = 0x0400;

            [UnmanagedFunctionPointer(CallingConvention.StdCall)]
            private delegate int GetIDsOfNamesDelegate(IntPtr dispatch, ref Guid riid,
                [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] names,
                int nameCount, int lcid, [Out] int[] dispIds);

            [UnmanagedFunctionPointer(CallingConvention.StdCall)]
            private delegate int InvokeDelegate(IntPtr dispatch, int dispId, ref Guid riid, int lcid, short flags,
                ref ComTypes.DISPPARAMS dispParams, IntPtr result, IntPtr excepInfo, IntPtr argErr);

            [DllImport("oleaut32.dll")]
            private static extern int VariantClear(IntPtr variant);

            public static bool TryGetDispId(IntPtr dispatch, string name)
            {
                return GetIDsOfNames(dispatch, name, out _) == 0;
            }

            public static IntPtr InvokeReturningDispatch(IntPtr dispatch, string name, params object[] arguments)
            {
                Marshal.ThrowExceptionForHR(GetIDsOfNames(dispatch, name, out var dispId));

                var argumentVariants = Marshal.AllocCoTaskMem(VariantSize * Math.Max(arguments.Length, 1));
                var resultVariant = Marshal.AllocCoTaskMem(VariantSize);
                try
                {
                    // IDispatch::Invoke expects the arguments in reverse order.
                    for (var i = 0; i < arguments.Length; i++)
                    {
                        Marshal.GetNativeVariantForObject(arguments[arguments.Length - 1 - i], argumentVariants + i * VariantSize);
                    }
                    Marshal.GetNativeVariantForObject(null, resultVariant);

                    var dispParams = new ComTypes.DISPPARAMS
                    {
                        cArgs = arguments.Length,
                        rgvarg = arguments.Length > 0 ? argumentVariants : IntPtr.Zero
                    };
                    var riid = Guid.Empty;
                    var invoke = GetVtableFunction<InvokeDelegate>(dispatch, InvokeVtableSlot);
                    Marshal.ThrowExceptionForHR(invoke(dispatch, dispId, ref riid, LocaleUserDefault,
                                                       DispatchMethodOrPropertyGet, ref dispParams,
                                                       resultVariant, IntPtr.Zero, IntPtr.Zero));

                    Assert.That(Marshal.ReadInt16(resultVariant), Is.EqualTo(VariantTypeDispatch), "Result is not VT_DISPATCH");
                    var resultDispatch = Marshal.ReadIntPtr(resultVariant, 8);
                    Marshal.AddRef(resultDispatch);
                    return resultDispatch;
                }
                finally
                {
                    for (var i = 0; i < arguments.Length; i++)
                    {
                        VariantClear(argumentVariants + i * VariantSize);
                    }
                    VariantClear(resultVariant);
                    Marshal.FreeCoTaskMem(argumentVariants);
                    Marshal.FreeCoTaskMem(resultVariant);
                }
            }

            private static int GetIDsOfNames(IntPtr dispatch, string name, out int dispId)
            {
                var riid = Guid.Empty;
                var dispIds = new int[1];
                var getIDsOfNames = GetVtableFunction<GetIDsOfNamesDelegate>(dispatch, GetIDsOfNamesVtableSlot);
                var hresult = getIDsOfNames(dispatch, ref riid, new[] { name }, 1, LocaleUserDefault, dispIds);
                dispId = dispIds[0];
                return hresult;
            }

            private static T GetVtableFunction<T>(IntPtr comInterface, int slot) where T : class
            {
                var vtable = Marshal.ReadIntPtr(comInterface);
                var function = Marshal.ReadIntPtr(vtable, slot * IntPtr.Size);
                return Marshal.GetDelegateForFunctionPointer<T>(function);
            }
        }
    }
}
