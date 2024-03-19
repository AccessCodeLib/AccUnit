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

Implements AccUnit.ITestResultReporter

Private WithEvents m_TestResultCollector As AccUnit.TestResultCollector
Attribute m_TestResultCollector.VB_VarHelpID = -1

Private Property Get ITestResultReporter_TestResultCollector() As AccUnit.ITestResultCollector
   Set ITestResultReporter_TestResultCollector = m_TestResultCollector
End Property

Private Property Set ITestResultReporter_TestResultCollector(ByVal TRC As AccUnit.ITestResultCollector)
   Set m_TestResultCollector = TRC
End Property

Private Sub m_TestResultCollector_NewTestResult(ByVal Result As AccUnit.ITestResult)
'
End Sub

Private Sub m_TestResultCollector_TestSuiteFinished(ByVal Summary As AccUnit.ITestSummary)
'
End Sub

Private Sub m_TestResultCollector_TestSuiteStarted(ByVal TestSuite As AccUnit.ITestSuite)
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

Private Sub m_TestResultCollector_PrintSummary(ByVal TestSummary As AccUnit.ITestSummary, ByVal PrintTestResults As Boolean)
   PrintSummary TestSummary, PrintTestResults
End Sub

Private Sub PrintSummary(ByVal TestSummary As AccUnit.ITestSummary, Optional ByVal PrintTestResults As Boolean = False)
   If PrintTestResults Then
      Debug.Print String(40, "-")
      PrintTestDetailSummary TestSummary.GetTestResults
   End If
   Debug.Print String(40, "-")
   Debug.Print "Tests: "; TestSummary.Total
   Debug.Print , "Passed:  "; TestSummary.Passed
   Debug.Print , "Failed:  "; TestSummary.Failed
   Debug.Print , "Ignored: "; TestSummary.Ignored
   Debug.Print String(40, "-")
End Sub

Private Sub PrintTestDetailSummary(ByRef TestResults() As AccUnit.ITestResult)
   
   Dim i As Long
   Dim r As Variant ' AccUnit.ITestResult
   Dim LastTestIsRowTest As Boolean
   
   For Each r In TestResults
      i = i + 1
      If TypeOf r Is AccUnit.TestResultCollection Then
         If i > 1 Then Debug.Print String(20, "-")
         Debug.Print CStr(i), r.Test.FullName & ":", r.Result, "..."
         Debug.Print String(3, " ") & String(17, "-")
         PrintSubResults i, r
         LastTestIsRowTest = True
      Else
         If LastTestIsRowTest Then
            LastTestIsRowTest = False
            If i > 1 Then Debug.Print String(20, "-")
         End If
         Debug.Print CStr(i), r.Test.FullName & ":", r.Result, r.Message
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
         Debug.Print String(3, " ") & ResultID, r.Test.FullName & "-", r.Result, "..."
         PrintSubResults ResultID, r
      Else
         Debug.Print String(3, " ") & ResultID, r.Test.FullName & "-", r.Result, r.Message
      End If
   Next

End Sub

