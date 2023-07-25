using AccessCodeLib.AccUnit.Common;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using VBA;
using VbMsgBoxResult = AccessCodeLib.AccUnit.Interfaces.VbMsgBoxResult;

namespace AccessCodeLib.AccUnit
{
    [ComVisible(true)]
    [Guid("3A27B2A4-99E9-4427-B0F0-D1F0053B90B2")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITestMessageBoxComEvents))]
    [ProgId("AccUnit.TestMessageBox")]
    public class TestMessageBox : ITestMessageBox
    {
        private const string SetTestMsgBoxProcedureName = "SetAccUnitTestMsgBox";

        public delegate void TestMessageBoxDisplayedEventHandler(object Prompt, int Buttons, object Title, object HelpFile, object Context, ref int MsgBoxResult);
        public event TestMessageBoxDisplayedEventHandler Displayed;

        private readonly IList<MessageBoxData> _messageBoxDataList = new List<MessageBoxData>();
        private readonly Queue<VbMsgBoxResult> _vbMsgBoxResults = new Queue<VbMsgBoxResult>();

        public MessageBoxData[] MessageBoxData { get { return _messageBoxDataList.ToArray(); } }

        public void InitMsgBoxResults(params VbMsgBoxResult[] args)
        {
            using (new BlockLogger())
            {
                try
                {
                    foreach (var vbMsgBoxResult in args)
                    {
                        Logger.Log(string.Format("insert {0}", vbMsgBoxResult));
                        _vbMsgBoxResults.Enqueue(vbMsgBoxResult);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
            }
        }

        public void InitMsgBoxResults(IList<VbMsgBoxResult> args)
        {
            foreach (var vbMsgBoxResult in args)
            {
                _vbMsgBoxResults.Enqueue(vbMsgBoxResult);
            }
        }

        internal VbMsgBoxResult NextMsgBoxResult
        {
            get
            {
                using (new BlockLogger())
                {

                    if ((_vbMsgBoxResults == null) || (_vbMsgBoxResults.Count == 0))
                        throw new MissingTestMessageBoxResultsException();

                    try
                    {
                        return _vbMsgBoxResults.Dequeue();
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(ex);
                        return 0;
                    }
                }
            }
        }

        [ComVisible(true)]
        public int Show(object prompt, int buttons = (int)VbMsgBoxStyle.vbOKOnly,
                                   object title = null, object helpFile = null, object context = null)
        {
            var msgBoxResult = (int)NextMsgBoxResult;
            Displayed?.Invoke(prompt, buttons, title, helpFile, context, ref msgBoxResult);

            _messageBoxDataList.Add(new MessageBoxData
            {
                Prompt = prompt,
                Buttons = buttons,
                Title = title,
                HelpFile = helpFile,
                Context = context,
                MsgBoxResult = msgBoxResult
            });
            return msgBoxResult;
        }


        public static void DisposeTestMessageBox(OfficeApplicationHelper officeApplicationHelper)
        {
            try
            {
                officeApplicationHelper.Run(SetTestMsgBoxProcedureName, null);
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                /* ignore error */
            }
        }

        public void ActivateTestMessageBox(OfficeApplicationHelper officeApplicationHelper, ITestMessageBox messageBox)
        {
            officeApplicationHelper.Run(SetTestMsgBoxProcedureName, messageBox);
        }

        public static void CheckTestMessageBoxProcedures(VBProject vbProject)
        {
            var testClassFactory = new TestClassFactoryManager(vbProject, new TestClassReader(vbProject));
            var module = testClassFactory.FactoryModule;
            if (!AccUnitTestMsgBoxPropertySetExists(module))
                InsertAccUnitTestMsgBoxProcedures(module);
        }

        private static bool AccUnitTestMsgBoxPropertySetExists(_CodeModule codeModule)
        {
            try
            {
                if (codeModule.ProcStartLine[SetTestMsgBoxProcedureName, vbext_ProcKind.vbext_pk_Proc] > 0)
                    return true;
            }
            catch (Exception ex)
            {
                /* ignore error */
                Logger.Log(ex);
            }
            return false;
        }

        private static void InsertAccUnitTestMsgBoxProcedures(_CodeModule codeModule)
        {
            var declarationLines = codeModule.CountOfDeclarationLines;
            var startLine = declarationLines + 1;
            codeModule.InsertLines(startLine, TestMessageBoxSource);
        }

        private static string TestMessageBoxSource
        {
            get
            {
                return @"
Private m_AccUnitTestMsgBox As AccUnit_Integration.TestMessageBox

Public Sub SetAccUnitTestMsgBox(ByRef NewRef As AccUnit_Integration.TestMessageBox)
   Set m_AccUnitTestMsgBox = NewRef
End Sub

Public Function MsgBox(ByVal Prompt As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                       Optional ByVal Title As Variant, Optional ByVal HelpFile As Variant, _
                       Optional ByVal Context As Variant) As VbMsgBoxResult

    If m_AccUnitTestMsgBox Is Nothing Then
        MsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
    Else
        MsgBox = m_AccUnitTestMsgBox.Show(Prompt, Buttons, Title, HelpFile, Context)
    End If

End Function";
            }
        }

        private static readonly Regex UsedInCodeModuleRegex = new Regex(@"^\s*\'\s*AccUnit:.*ClickingMsgBox\(.*$", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        public static bool UsedInCodeModule(_CodeModule codeModule)
        {
            return UsedInCodeModuleRegex.IsMatch(codeModule.Lines[1, codeModule.CountOfLines]);
        }
    }

    [ComVisible(true)]
    [Guid("3D29E65A-3A3A-4E43-A1C1-9EBEBF73E898")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITestMessageBoxComEvents
    {
        void Displayed(object Prompt, int Buttons, object Title, object HelpFile, object Context, ref int MsgBoxResult);
    }
}