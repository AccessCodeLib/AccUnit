VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugPrintTestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Klassenmodul: DebugPrintTestSuite
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Debug.Print-Ausgabe vom Test-Ergebnis
' </summary>
' <remarks>
' </remarks>
'\ingroup addins_shared
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/AccUnitLoader/DebugPrintTestSuite.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Implements AccUnit.IVBATestSuite

Private WithEvents m_TestSuite As AccUnit.VbaTestSuite
Attribute m_TestSuite.VB_VarHelpID = -1
Private m_TestResultCollector As DebugPrintTestResultCollector
Private m_CodeCoverageTracker As AccUnit.CodeCoverageTracker
'

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
   m_TestSuite.Dispose
   Set m_TestSuite = Nothing
End Sub

Public Sub Init(ByVal NewTestSuite As AccUnit.VbaTestSuite, _
                ByVal NewTestResultCollector As DebugPrintTestResultCollector)
   Set m_TestSuite = NewTestSuite
   Set m_TestResultCollector = NewTestResultCollector
   Set m_TestSuite.TestResultCollector = m_TestResultCollector
End Sub

Public Property Get TestSuite() As AccUnit.VbaTestSuite
   Set TestSuite = m_TestSuite
End Property

Private Sub m_TestSuite_TestTraceMessage(ByVal Message As String)
   If Message Like "* [*][*][*]*" Then
      Message = Replace(Message, vbNewLine, vbTab)
   End If
   Debug.Print Message
   
   If Not (m_CodeCoverageTracker Is Nothing) Then
   If Message Like "*Total*" Then
      Debug.Print m_CodeCoverageTracker.GetReport
   End If
   End If
End Sub

Public Sub PrintSummary(Optional ByVal PrintTestResults As Boolean = False)
   m_TestResultCollector.PrintSummary PrintTestResults
End Sub

Public Property Set CodeCoverageTracker(ByVal NewRef As Object)
   Set m_CodeCoverageTracker = NewRef
End Property

'####################################
' IVBATestSuite
'
Private Function IVBATestSuite_Reset(Optional ByVal mode As AccUnit.ResetMode = 1&) As AccUnit.IVBATestSuite
   m_TestSuite.Reset mode
   Set IVBATestSuite_Reset = Me
End Function

Private Property Set IVBATestSuite_ActiveVBProject(ByVal RHS As VBIDE.VBProject)
   Set m_TestSuite.ActiveVBProject = RHS
End Property

Private Property Get IVBATestSuite_ActiveVBProject() As VBIDE.VBProject
   Set IVBATestSuite_ActiveVBProject = m_TestSuite.ActiveVBProject
End Property

Public Function Add(ByVal testToAdd As Object) As DebugPrintTestSuite
   m_TestSuite.Add testToAdd
   Set Add = Me
End Function

Private Function IVBATestSuite_Add(ByVal testToAdd As Object) As AccUnit.IVBATestSuite
   Set IVBATestSuite_Add = Add(testToAdd)
End Function

Public Function AddByClassName(ByVal className As String) As DebugPrintTestSuite
   m_TestSuite.AddByClassName className
   Set AddByClassName = Me
End Function

Private Function IVBATestSuite_AddByClassName(ByVal className As String) As AccUnit.IVBATestSuite
   Set IVBATestSuite_AddByClassName = AddByClassName(className)
End Function

Public Function AddFromVBProject() As DebugPrintTestSuite
   m_TestSuite.AddFromVBProject
   Set AddFromVBProject = Me
End Function

Private Function IVBATestSuite_AddFromVBProject() As AccUnit.IVBATestSuite
   Set IVBATestSuite_AddFromVBProject = AddFromVBProject()
End Function

Private Sub IVBATestSuite_Dispose()
   m_TestSuite.Dispose
End Sub

Private Property Set IVBATestSuite_HostApplication(ByVal RHS As Object)
   Set m_TestSuite.HostApplication = RHS
End Property

Private Property Get IVBATestSuite_HostApplication() As Object
   Set IVBATestSuite_HostApplication = m_TestSuite.HostApplication
End Property

Private Property Get IVBATestSuite_Name() As String
   IVBATestSuite_Name = "DebugPrintTestSuite"
End Property

Public Function Run() As DebugPrintTestSuite
   m_TestSuite.Run
   Set Run = Me
End Function

Private Function IVBATestSuite_Run() As AccUnit.IVBATestSuite
   Set IVBATestSuite_Run = Run()
End Function

Private Property Get IVBATestSuite_Summary() As AccUnit.ITestSummary
   Set IVBATestSuite_Summary = m_TestSuite.Summary
End Property

Private Property Set IVBATestSuite_TestResultCollector(ByVal RHS As AccUnit.ITestResultCollector)
   Set m_TestSuite.TestResultCollector = RHS
End Property

Private Property Get IVBATestSuite_TestResultCollector() As AccUnit.ITestResultCollector
   Set IVBATestSuite_TestResultCollector = m_TestSuite.TestResultCollector
End Property

Private Property Set IVBATestSuite_TestRunner(ByVal RHS As AccUnit.ITestRunner)
   Set m_TestSuite.TestRunner = RHS
End Property

Private Property Get IVBATestSuite_TestRunner() As AccUnit.ITestRunner
   Set IVBATestSuite_TestRunner = m_TestSuite.TestRunner
End Property
