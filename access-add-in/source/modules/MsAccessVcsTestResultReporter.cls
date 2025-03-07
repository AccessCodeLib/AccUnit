﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MsAccessVcsTestResultReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: LogFileTestResultReporter
'---------------------------------------------------------------------------------------
'
' Output test results to log file (text file)
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private Const LineSeperatorLen As Long = 40
Private Const MsAccessVcsAddInFileName As String = "Version Control"

Implements AccUnit.ITestResultReporter
Private m_VCS As Object

Private WithEvents m_TestResultCollector As AccUnit.TestResultCollector
Attribute m_TestResultCollector.VB_VarHelpID = -1

Private Sub Class_Terminate()
   Set m_VCS = Nothing
End Sub

Private Property Get ITestResultReporter_TestResultCollector() As AccUnit.ITestResultCollector
   Set ITestResultReporter_TestResultCollector = m_TestResultCollector
End Property

Private Property Set ITestResultReporter_TestResultCollector(ByVal TestResultCollectorToListen As ITestResultCollector)
   Set TestResultCollector = TestResultCollectorToListen
End Property

Private Property Set TestResultCollector(ByVal TestResultCollectorToListen As TestResultCollector)
   Set m_TestResultCollector = TestResultCollectorToListen
End Property

Private Sub m_TestResultCollector_TestFinished(ByVal Result As AccUnit.ITestResult)
'
End Sub

Private Sub m_TestResultCollector_TestFixtureFinished(ByVal Result As AccUnit.ITestResult)
'
End Sub

Private Sub m_TestResultCollector_TestFixtureStarted(ByVal TestFixture As AccUnit.ITestFixture)
'
End Sub

Private Sub m_TestResultCollector_TestStarted(ByVal Test As AccUnit.ITest, ByVal IgnoreInfo As AccUnit.IgnoreInfo)
'
End Sub

Private Sub m_TestResultCollector_TestSuiteReset(ByVal Summary As AccUnit.ITestSummary)
'
End Sub

Private Sub m_TestResultCollector_TestSuiteStarted(ByVal TestSuite As AccUnit.ITestSuite)
   PrintToFile String(LineSeperatorLen, "#")
   PrintToFile Format(Now(), "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub m_TestResultCollector_TestTraceMessage(ByVal Message As String, ByVal CodeCoverageTracker As AccUnit.ICodeCoverageTracker)
'
End Sub

Private Sub m_TestResultCollector_NewTestResult(ByVal Result As AccUnit.ITestResult)
'
End Sub

Private Sub m_TestResultCollector_TestSuiteFinished(ByVal TestSummary As AccUnit.ITestSummary)
   PrintSummary TestSummary, True
   PrintToFile String(LineSeperatorLen, "#")
End Sub

Private Sub m_TestResultCollector_PrintSummary(ByVal TestSummary As AccUnit.ITestSummary, ByVal PrintTestResults As Boolean)
   PrintSummary TestSummary, PrintTestResults
End Sub

Private Sub PrintSummary(ByVal TestSummary As AccUnit.ITestSummary, ByVal PrintTestResults As Boolean)

   If PrintTestResults Then
      PrintToFile String(LineSeperatorLen, "-")
      PrintTestDetailSummary TestSummary.GetTestResults
   End If
   PrintToFile String(LineSeperatorLen, "-")
   PrintToFile "Tests: " & TestSummary.Total
   PrintToFile " ", "Passed:  " & TestSummary.Passed
   PrintToFile " ", "Failed:  " & TestSummary.Failed
   PrintToFile " ", "Error:   " & TestSummary.Error
   PrintToFile " ", "Ignored: " & TestSummary.Ignored
   PrintToFile String(LineSeperatorLen, "-")

   If TestSummary.Failed + TestSummary.Error > 0 Then
      PrintToFile (TestSummary.Failed + TestSummary.Error) & " / " & TestSummary.Total & " failed"
   ElseIf TestSummary.Passed = TestSummary.Total Then
      PrintToFile (TestSummary.Passed) & " / " & TestSummary.Total & " passed"
   Else
      PrintToFile (TestSummary.Ignored) & " / " & TestSummary.Total & " ignored"
   End If

   PrintToFile String(LineSeperatorLen, "-")

End Sub

Private Sub PrintTestDetailSummary(ByRef TestResults() As AccUnit.ITestResult)

   Dim i As Long
   Dim r As Variant 'AccUnit.ITestResult
   Dim LastTestIsRowTest As Boolean

   For Each r In TestResults
      i = i + 1
      If TypeOf r Is AccUnit.TestResultCollection Then
         If i > 1 Then PrintToFile String(20, "-")
         PrintToFile CStr(i), r.Test.FullName & ":", r.Result, "..."
         PrintToFile String(3, " ") & String(17, "-")
         PrintSubResults i, r
         LastTestIsRowTest = True
      Else
         If LastTestIsRowTest Then
            LastTestIsRowTest = False
            If i > 1 Then PrintToFile String(20, "-")
         End If
         PrintToFile CStr(i), r.Test.FullName & ":", r.Result, r.Message
      End If
   Next

End Sub

Private Sub PrintSubResults(ByVal MainId As String, ByVal ResultCol As AccUnit.TestResultCollection)

   Dim i As Long
   Dim r As AccUnit.ITestResult
   Dim ResultID As String

   For i = 1 To ResultCol.Count
      Set r = ResultCol.Item(i - 1)
      ResultID = MainId & "." & i
      If TypeOf r Is AccUnit.ITestResultSummary Then
         PrintToFile String(3, " ") & ResultID, r.Test.FullName, r.Result, "..."
         PrintSubResults ResultID, r
      Else
         PrintToFile String(3, " ") & ResultID, r.Test.FullName, r.Result, r.Message
      End If
   Next

End Sub

Private Sub PrintToFile(ParamArray LogItems())

   Dim LineString As String
   Dim i As Long

   LineString = LogItems(LBound(LogItems))
   For i = LBound(LogItems) + 1 To UBound(LogItems)
      LineString = LineString & vbTab & LogItems(i)
   Next

   If m_VCS Is Nothing Then
      InitVcs
   End If

'Interface:
'Public Sub AddLog(strText As String, Optional blnPrint As Boolean = True, _
'    Optional blnNextOutputOnNewLine As Boolean = True, _
'    Optional strColor As String = vbNullString, _
'    Optional blnBold As Boolean = False, _
'    Optional blnItalic As Boolean = False)
'
   m_VCS.AddLog LineString, False

End Sub

Private Sub InitVcs()
   Set m_VCS = Application.Run(GetVcsAddInRunPath & ".VCS")
End Sub

Private Function GetVcsAddInRunPath() As String

   Dim AddInPath As String
   AddInPath = Environ("appdata") & "\MsAccessVCS\"

   GetVcsAddInRunPath = AddInPath & MsAccessVcsAddInFileName

End Function
