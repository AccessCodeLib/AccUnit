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

Private Const LineSeperatorLen As Long = 40

Implements ITestResultReporter
Private m_FileNo As Long

Private WithEvents m_TestResultCollector As TestResultCollector
Attribute m_TestResultCollector.VB_VarHelpID = -1

Private Sub Class_Terminate()
   CloseFile
End Sub

Private Property Set ITestResultReporter_TestResultCollector(ByVal TestResultCollectorToListen As TestResultCollector)
   Set TestResultCollector = TestResultCollectorToListen
End Property

Private Property Set TestResultCollector(ByVal TestResultCollectorToListen As TestResultCollector)
   Set m_TestResultCollector = TestResultCollectorToListen
End Property

Private Sub m_TestResultCollector_TestSuiteFinished(ByVal TestSummary As AccUnit.ITestSummary)
   PrintSummary TestSummary, True
   CloseFile
End Sub

Private Sub m_TestResultCollector_TestSuiteStarted(ByVal TestSuite As AccUnit.ITestSuite)
   PrintToFile String(LineSeperatorLen, "#")
   PrintToFile Format(Now(), "yyyy-mm-dd hh:nn:ss")
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
         PrintToFile CStr(i), r.test.FullName & ":", r.Result, "..."
         PrintToFile String(3, " ") & String(17, "-")
         PrintSubResults i, r
         LastTestIsRowTest = True
      Else
         If LastTestIsRowTest Then
            LastTestIsRowTest = False
            If i > 1 Then PrintToFile String(20, "-")
         End If
         PrintToFile CStr(i), r.test.FullName & ":", r.Result, r.Message
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
         PrintToFile String(3, " ") & ResultID, r.test.FullName, r.Result, "..."
         PrintSubResults ResultID, r
      Else
         PrintToFile String(3, " ") & ResultID, r.test.FullName, r.Result, r.Message
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

   If m_FileNo = 0 Then
      OpenFile
   End If

   Print #m_FileNo, LineString

End Sub

Private Sub OpenFile()
   m_FileNo = FreeFile
   Open LogFilePath For Append As #m_FileNo
End Sub

Private Sub CloseFile()
On Error Resume Next
   If m_FileNo <> 0 Then
      Close #m_FileNo
      m_FileNo = 0
   End If
End Sub

Private Property Get LogFilePath() As String
   LogFilePath = CurrentVbProject.FileName & ".AccUnit.log"
End Property
