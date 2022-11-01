using AccessCodeLib.Common.VBIDETools.Templates;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.Configuration
{
    public class TestSuiteCodeTemplates : CodeTemplateCollection
    {
        public TestSuiteCodeTemplates()
        {
            Add(TestSuiteStarter);
            Add(AccUnitObjects);
        }

        private static CodeTemplate TestSuiteStarter => new CodeTemplate(@"TestSuiteStarter", vbext_ComponentType.vbext_ct_ClassModule,
            @"Option Compare Text
Option Explicit

Private WithEvents m_TestSuite As AccUnit_Integration.VBATestSuite
Private WithEvents m_TestSuiteDebugOutput As AccUnit_Integration.VBATestSuite

Private Sub InitTestSuite()
   Set m_TestSuiteDebugOutput = Nothing
   Set m_TestSuite = GetTestSuiteFromAddIn

   If m_TestSuite Is Nothing Then
      If Application.Name = ""Microsoft Access"" Then
         Dim accSuite As AccUnit_Integration.AccessTestSuite
         Set accSuite = New AccUnit_Integration.AccessTestSuite
         Set accSuite.HostApplication = Application
         Set m_TestSuite = accSuite
      Else
         Set m_TestSuite = New AccUnit_Integration.VBATestSuite
         Set m_TestSuite.HostApplication = Application
         Set m_TestSuite.ActiveVBProject = Application.VBE.ActiveVBProject
      End If
      Set m_TestSuiteDebugOutput = m_TestSuite
   End If
End Sub

Private Function GetTestSuiteFromAddIn() As AccUnit_Integration.VBATestSuite
   Dim TempAddIn As Object
   Dim TestSuitefromAddIn As AccUnit_Integration.VBATestSuite
   Dim AddInManger As Object
   
   For Each TempAddIn In Application.VBE.AddIns
      If TempAddIn.progID = ""AccUnit.GUI.Connect"" Then
         Set AddInManger = TempAddIn.Object
         If Not (AddInManger Is Nothing) Then
            Set AddInManger.Application = Application
            Set TestSuitefromAddIn = AddInManger.TestSuite
         End If
         Exit For
      End If
   Next
   
   Set GetTestSuiteFromAddIn = TestSuitefromAddIn
End Function

Public Property Get AccUnitTestSuite() As AccUnit_Integration.VBATestSuite
   If m_TestSuite Is Nothing Then
      initTestSuite
   End If
   Set AccUnitTestSuite = m_TestSuite
End Property

Private Sub m_TestSuite_Disposed(ByVal sender As Object)
   Set m_TestSuite = Nothing
   Set m_TestSuiteDebugOutput = Nothing
End Sub

Private Sub m_TestSuite_TestTraceMessage(ByVal Message As String)
'  Debug.Print Message
End Sub

Private Sub m_TestSuiteDebugOutput_TestTraceMessage(ByVal Message As String)
   Debug.Print Message
End Sub");

        private static CodeTemplate AccUnitObjects => new CodeTemplate("AccUnitObjects", vbext_ComponentType.vbext_ct_StdModule,
            @"Option Compare Text
Option Explicit

Private m_TestSuiteStarter As TestSuiteStarter

Public Property Get TestSuite() As AccUnit_Integration.VBATestSuite
    If m_TestSuiteStarter Is Nothing Then
        Set m_TestSuiteStarter = New TestSuiteStarter
    End If
    Set TestSuite = m_TestSuiteStarter.AccUnitTestSuite
End Property");
    }
}