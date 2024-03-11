VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogFileTestResultReporter"
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

Private Sub m_TestResultCollector_TestSuiteFinished(ByVal TestSummary As AccUnit.ITestSummary)
   PrintSummary TestSummary, True
   PrintToFile String(40, "#")
End Sub

Private Sub m_TestResultCollector_TestSuiteStarted(ByVal TestSuite As AccUnit.ITestSuite)
   PrintToFile String(40, "#")
End Sub

Private Sub m_TestResultCollector_TestTraceMessage(ByVal Message As String, ByVal CodeCoverageTracker As CodeCoverageTracker)
'   If Message Like "* [*][*][*]*" Then
'      Message = Replace(Message, vbNewLine, vbTab)
'   ElseIf Message Like "* ~~~*" Then
'      Message = Replace(Message, vbNewLine, vbTab)
'   End If
'   PrintToFile Message
'
'   If Not (CodeCoverageTracker Is Nothing) Then
'   If Message Like "*Total*" Then
'      PrintToFile CodeCoverageTracker.GetReport()
'   End If
'   End If
End Sub

Private Sub m_TestResultCollector_PrintSummary(ByVal TestSummary As AccUnit.ITestSummary, ByVal PrintTestResults As Boolean)
   PrintSummary TestSummary, PrintTestResults
End Sub

Private Sub PrintSummary(ByVal TestSummary As AccUnit.ITestSummary, ByVal PrintTestResults As Boolean)
   If PrintTestResults Then
      PrintToFile String(40, "-")
      PrintTestDetailSummary TestSummary.GetTestResults
   End If
   PrintToFile String(40, "-")
   PrintToFile "Tests: " & TestSummary.Total
   PrintToFile " ", "Passed:  " & TestSummary.Passed
   PrintToFile " ", "Failed:  " & TestSummary.Failed
   PrintToFile " ", "Ignored: " & TestSummary.Ignored
   PrintToFile String(40, "-")
   
   If TestSummary.Failed + TestSummary.Error > 0 Then
      PrintToFile (TestSummary.Failed + TestSummary.Error) & " / " & TestSummary.Total & " failed"
   ElseIf TestSummary.Passed = TestSummary.Total Then
      PrintToFile (TestSummary.Passed) & " / " & TestSummary.Total & " passed"
   Else
      PrintToFile (TestSummary.Ignored) & " / " & TestSummary.Total & " ignored"
   End If
   
   PrintToFile String(40, "-")
   
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

Private Sub PrintSubResults(ByVal mainId As String, ByVal resultCol As AccUnit.TestResultCollection)
   
   Dim i As Long
   Dim r As AccUnit.ITestResult
   Dim ResultID As String
   
   For i = 1 To resultCol.Count
      Set r = resultCol.Item(i - 1)
      ResultID = mainId & "." & i
      If TypeOf r Is AccUnit.ITestResultSummary Then
         PrintToFile String(3, " ") & ResultID, r.Test.FullName, r.Result, "..."
         PrintSubResults ResultID, r
      Else
         PrintToFile String(3, " ") & ResultID, r.Test.FullName, r.Result, r.Message
      End If
   Next

End Sub

Private Sub PrintToFile(ParamArray LogItems())
   
   Dim FileNo As Long
   FileNo = FreeFile
   
   Dim LineString As String
   Dim i As Long
   
   LineString = LogItems(LBound(LogItems))
   For i = LBound(LogItems) + 1 To UBound(LogItems)
      LineString = LineString & vbTab & LogItems(i)
   Next
      
   Open LogFilePath For Append As #FileNo
   Print #FileNo, LineString
   Close #FileNo
   
End Sub

Private Property Get LogFilePath() As String
   LogFilePath = Application.ActiveWorkbook.FullName & ".AccUnit.log"
End Property


