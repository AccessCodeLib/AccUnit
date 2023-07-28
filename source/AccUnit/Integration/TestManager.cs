using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Integration;
using Microsoft.Vbe.Interop;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Exception = System.Exception;
using ITest = AccessCodeLib.AccUnit.Interfaces.ITest;
using VbMsgBoxResult = AccessCodeLib.AccUnit.Interfaces.VbMsgBoxResult;

namespace AccessCodeLib.AccUnit
{
    [ComVisible(true)]
    [Guid("A1CB378B-9DAB-4652-85D1-268E7A2C9AA7")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.TestManager")]
    public class TestManager : ITestManagerComInterface
    {
        public const string RowNameDelimiter = TestRow.RowNameDelimiter;

        private readonly ITestManagerBridge _testObject;
        private readonly TestClassMemberList _memberFilter;
        private readonly TestRowGenerator _testRowGenerator = new TestRowGenerator();

        private readonly TestClassMemberList _members = new TestClassMemberList();
        private OfficeApplicationHelper _officeApplicationHelper;

        public TestManager()
        {
        }

        public TestManager(ITestManagerBridge testmanagerbridge, TestClassMemberList memberFilter = null)
            : this()
        {
            try
            {
                TestName = Information.TypeName(testmanagerbridge);
            }
            catch (Exception ex) { Logger.Log(ex.Message); }

            _memberFilter = memberFilter;
            _testObject = testmanagerbridge;
            _testObject.InitTestManager(this);
        }

        private string _TestName;
        public string TestName
        {
            get { return _TestName; }
            set
            {
                _TestName = value;
                _testRowGenerator.TestName = value;
            }
        }

        public object HostApplication
        {
            get { return _officeApplicationHelper?.Application; }
            set
            {
                _officeApplicationHelper = ComTools.GetTypeForComObject(value, "Access.Application") != null
                                                ? new AccessApplicationHelper(value) : new OfficeApplicationHelper(value);
            }
        }

        private VBProject _activeVBProject;
        public VBProject ActiveVBProject
        {
            get { return _activeVBProject; }
            set
            {
                _activeVBProject = value;
                _testRowGenerator.ActiveVBProject = value;
            }
        }

        public void AddTests(ITestFixture testFixture)
        {
            throw new NotImplementedException("TODO: replace TypeLibTools.GetTLIInterfaceMembers (TypleLib less) ");
            // TODO: replace TypeLibTools.GetTLIInterfaceMembers (TypleLib less) 
            /*
            foreach (TLI.MemberInfo member in TypeLibTools.GetTLIInterfaceMembers(_testObject))
            {
                
                if (IsSetupOrTeardown(member, testFixture))
                {
                    continue;
                }

                var name = member.Name;
                var useGetTestData = (member.Parameters.Count > 0);
                TestClassMemberInfo memberinfo;

                if (_memberFilter != null)
                {
                    memberinfo = FindTestClassMemberInfo(name);
                    if (memberinfo == null)
                    {
                        continue;
                    }
                }
                else
                {
                    memberinfo = new TestClassMemberInfo(member.Name);
                }

            

               
                //var databuilder = testCaseCollector.Add(name);
                //if (useGetTestData)
                //{
                //    _testRowGenerator.GetTestData(databuilder, memberinfo);
                //}
                

                _members.Add(memberinfo);

                // init TestMessagebox
               
            }
            */
        }

        public TestClassMemberList Members { get { return _members; } }

        private static bool IsSetupOrTeardown(string memberName, ITestFixture testFixtureInfo)
        {
            switch (memberName.ToLower())
            {
                case "fixturesetup":
                    testFixtureInfo.HasFixtureSetup = true;
                    break;
                case "setup":
                    testFixtureInfo.HasSetup = true;
                    break;
                case "teardown":
                    testFixtureInfo.HasTeardown = true;
                    break;
                case "fixtureteardown":
                    testFixtureInfo.HasFixtureTeardown = true;
                    break;
                default:
                    return false;
            }
            return true;
        }

        private TestClassMemberInfo FindTestClassMemberInfo(string name)
        {
            return _memberFilter.Find(
                m => m.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase)
                );
        }

        #region row test data

        public _ITestRow Row(object arg1 = null, object arg2 = null, object arg3 = null, object arg4 = null, object arg5 = null,
                        object arg6 = null, object arg7 = null, object arg8 = null, object arg9 = null, object arg10 = null,
                        object arg11 = null, object arg12 = null, object arg13 = null, object arg14 = null, object arg15 = null)
        {
            throw new NotSupportedException("This method is only for IntelliSense in VBA editor.");
        }

        /*
        public void GetTestData(TestDataBuilder databuilder)
        {
            _testRowGenerator.GetTestData(databuilder);
        }
        */

        #endregion

        public ITestMessageBox TestMessageBox { get; private set; }

        public _ITestRow ClickingMsgBox(VbMsgBoxResult arg1 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg2 = VbMsgBoxResult.vbOK,
                                        VbMsgBoxResult arg3 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg4 = VbMsgBoxResult.vbOK,
                                        VbMsgBoxResult arg5 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg6 = VbMsgBoxResult.vbOK,
                                        VbMsgBoxResult arg7 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg8 = VbMsgBoxResult.vbOK,
                                        VbMsgBoxResult arg9 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg10 = VbMsgBoxResult.vbOK,
                                        VbMsgBoxResult arg11 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg12 = VbMsgBoxResult.vbOK,
                                        VbMsgBoxResult arg13 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg14 = VbMsgBoxResult.vbOK,
                                        VbMsgBoxResult arg15 = VbMsgBoxResult.vbOK)
        {
            throw new NotSupportedException("This method is only for IntelliSense in VBA editor.");
        }

        public void InitTestMessageBox(ITest testcase)
        {
            using (new BlockLogger("InitTestMessageBox"))
            {

                if (_officeApplicationHelper is null)
                    throw new NullReferenceException("OfficeApplicationHelper");

                try
                {
                    TestMessageBox = null;

                    var test = GetTestFixtureFromTest(testcase);
                    //Logger.Log(string.Format("Test: {0}\nMethod: {1}\nTestCase: {2}", test.Name, testcase.MethodName, testcase.Name));
                    var member = _members.Find(m => m.Name == testcase.MethodName);
                    using (new BlockLogger())
                    {
                        if (member.TestRows.Count > 0)
                            TestMessageBox = GetTestMessageBox(member.TestRows, testcase.Name);
                    }

                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }

                if (TestMessageBox is null)
                    TestMessageBox = GetTestMessageBoxFromMethod(testcase.MethodName);

                if (TestMessageBox is null)
                    return;

                TestMessageBox.ActivateTestMessageBox(_officeApplicationHelper, TestMessageBox);
            }
        }

        private static ITestMessageBox GetTestMessageBox(IList<ITestRow> rows, string name)
        {
            using (new BlockLogger())
            {
                if (name.Length <= 3 || name.Substring(0, 3) != "Row")
                    return null;
                var rowname = name.Substring(3);
                var charIndex = rowname.IndexOf(RowNameDelimiter);
                if (charIndex > 0)
                    rowname = rowname.Substring(0, charIndex);
                Logger.Log(rowname);
                var index = Convert.ToInt32(rowname) - 1;
                var row = rows[index];
                return row.TestMessageBox;
            }
        }

        private TestMessageBox GetTestMessageBoxFromMethod(string name)
        {
            var reader = new TestClassReader(ActiveVBProject);
            Logger.Log(string.Format("TestName: {0}, Member: {1}", TestName, name));
            var memberInfo = reader.GetTestClassMemberInfo(TestName, name);
            var results = memberInfo.MsgBoxResults;

            if (results is null || results.Count == 0)
            {
                Logger.Log("0 MsgBoxResults");
                return null;
            }

            var msgbox = new TestMessageBox();
            msgbox.InitMsgBoxResults(results);

            return msgbox;
        }

        private static ITestFixture GetTestFixtureFromTest(ITest testcase)
        {
            return testcase.Fixture;
        }


        private static bool IsRowTest(ITest test, ITest testcase)
        {
            return test.Name.Equals(testcase.MethodName, StringComparison.CurrentCultureIgnoreCase);
        }

    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("24AF805F-13AB-4137-A252-A804B815BDEA")]
    public interface ITestManagerComInterface
    {
        string TestName { get; set; }
        //void AddTest(TestCollector Test);
        //void GetTestData(TestDataBuilder Test);

        _ITestRow Row(object arg1 = null, object arg2 = null, object arg3 = null, object arg4 = null, object arg5 = null,
                      object arg6 = null, object arg7 = null, object arg8 = null, object arg9 = null, object arg10 = null,
                      object arg11 = null, object arg12 = null, object arg13 = null, object arg14 = null, object arg15 = null);

        ITestMessageBox TestMessageBox { get; }

        _ITestRow ClickingMsgBox(VbMsgBoxResult arg1 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg2 = VbMsgBoxResult.vbOK,
                                 VbMsgBoxResult arg3 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg4 = VbMsgBoxResult.vbOK,
                                 VbMsgBoxResult arg5 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg6 = VbMsgBoxResult.vbOK,
                                 VbMsgBoxResult arg7 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg8 = VbMsgBoxResult.vbOK,
                                 VbMsgBoxResult arg9 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg10 = VbMsgBoxResult.vbOK,
                                 VbMsgBoxResult arg11 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg12 = VbMsgBoxResult.vbOK,
                                 VbMsgBoxResult arg13 = VbMsgBoxResult.vbOK, VbMsgBoxResult arg14 = VbMsgBoxResult.vbOK,
                                 VbMsgBoxResult arg15 = VbMsgBoxResult.vbOK);
    }
}
