VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugPrintTestResultReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: DebugPrintTestResultReporter
'---------------------------------------------------------------------------------------
'
' Output test results with Debug.Print
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Implements ITestResultReporter

Private WithEvents m_TestResultCollector As TestResultCollector
Attribute m_TestResultCollector.VB_VarHelpID = -1

Private Property Set ITestResultReporter_TestResultCollector(ByVal TestResultCollectorToListen As TestResultCollector)
   Set TestResultCollector = TestResultCollectorToListen
End Property

Private Property Set TestResultCollector(ByVal TestResultCollectorToListen As TestResultCollector)
   Set m_TestResultCollector = TestResultCollectorToListen
End Property

Private Sub m_TestResultCollector_NewTestResult(ByVal TestResult As AccUnit.ITestResult)
'
End Sub

Private Sub m_TestResultCollector_TestTraceMessage(ByVal Message As String, ByVal CodeCoverageTracker As CodeCoverageTracker)
   If Message Like "* [*][*][*]*" Then
      Message = Replace(Message, vbNewLine, vbTab)
   ElseIf Message Like "* ~~~*" Then
      Message = Replace(Message, vbNewLine, vbTab)
   End If
   Debug.Print Message

   If Not (CodeCoverageTracker Is Nothing) Then
   If Message Like "*Total*" Then
      Debug.Print CodeCoverageTracker.GetReport()
   End If
   End If
End Sub

Private Sub m_TestResultCollector_PrintSummary(ByVal PrintTestResults As Boolean, ByVal TestSummary As TestSummary, ByVal TestResults As Collection)
   PrintSummary PrintTestResults, TestSummary, TestResults
End Sub

Private Sub PrintSummary(ByVal PrintTestResults As Boolean, ByVal TestSummary As TestSummary, ByVal TestResults As Collection)
   If PrintTestResults Then
      Debug.Print String(40, "-")
      PrintTestDetailSummary TestResults
   End If
   Debug.Print String(40, "-")
   Debug.Print "Tests: "; TestSummary.Total
   Debug.Print , "Passed:  "; TestSummary.Passed
   Debug.Print , "Failed:  "; TestSummary.Failed
   Debug.Print , "Ignored2: "; TestSummary.Ignored
   Debug.Print String(40, "-")
End Sub

Private Sub PrintTestDetailSummary(ByVal TestResults As Collection)

   Dim i As Long
   Dim r As AccUnit.ITestResult
   Dim LastTestIsRowTest As Boolean

   For Each r In TestResults
      i = i + 1
      If TypeOf r Is AccUnit.TestResultCollection Then
         If i > 1 Then Debug.Print String(20, "-")
         Debug.Print CStr(i), r.test.FullName & ":", r.Result, "..."
         Debug.Print String(3, " ") & String(17, "-")
         PrintSubResults i, r
         LastTestIsRowTest = True
      Else
         If LastTestIsRowTest Then
            LastTestIsRowTest = False
            If i > 1 Then Debug.Print String(20, "-")
         End If
         Debug.Print CStr(i), r.test.FullName & ":", r.Result, r.Message
      End If
   Next

End Sub

Private Sub PrintSubResults(ByVal mainId As String, ByVal resultCol As AccUnit.TestResultCollection)

   Dim i As Long
   Dim r As AccUnit.ITestResult
   Dim ResultID As String

   For i = 1 To resultCol.Count
      Set r = resultCol.Item(i - 1)
      ResultID = mainId & "." & i
      If TypeOf r Is AccUnit.ITestResultSummary Then
         Debug.Print String(3, " ") & ResultID, r.test.FullName & "-", r.Result, "..."
         PrintSubResults ResultID, r
      Else
         Debug.Print String(3, " ") & ResultID, r.test.FullName & "-", r.Result, r.Message
      End If
   Next

End Sub
