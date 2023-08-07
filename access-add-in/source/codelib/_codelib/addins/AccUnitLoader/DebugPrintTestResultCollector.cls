VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugPrintTestResultCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Klassenmodul: DebugPrintTestResultCollector
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
'  <file>_codelib/addins/AccUnitLoader/DebugPrintTestResultCollector.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Implements AccUnit.ITestResultCollector
Implements AccUnit.ITestSummaryPrinter

Private m_TestResults As Collection
Private m_TestSummary As TestSummary

Private Sub Class_Initialize()
   Set m_TestResults = New Collection
   Set m_TestSummary = New TestSummary
End Sub

Private Sub ITestResultCollector_Add(ByVal TestResult As AccUnit.ITestResult)
   m_TestResults.Add TestResult
   m_TestSummary.SaveResult TestResult
End Sub

Public Sub PrintSummary(Optional ByVal PrintTestResults As Boolean = False)
   If PrintTestResults Then
      Debug.Print String(40, "-")
      PrintTestDetailSummary
   End If
   Debug.Print String(40, "-")
   Debug.Print "Tests: "; m_TestSummary.Total
   Debug.Print , "Passed:  "; m_TestSummary.Passed
   Debug.Print , "Failed:  "; m_TestSummary.Failed
   Debug.Print , "Ignored2: "; m_TestSummary.Ignored
   Debug.Print String(40, "-")
End Sub

Private Property Get ITestSummaryPrinter_Name() As String
   ITestSummaryPrinter_Name = "DebugPrintTestResultCollector"
End Property

Private Sub ITestSummaryPrinter_PrintSummary(Optional ByVal PrintTestResults As Boolean = False)
   PrintSummary PrintTestResults
End Sub

Private Sub PrintTestDetailSummary()
   
   Dim i As Long
   Dim r As AccUnit.ITestResult
   Dim LastTestIsRowTest As Boolean
   
   For Each r In m_TestResults
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

Public Property Get Summary() As TestSummary
   Set Summary = m_TestSummary
End Property

Public Property Get TestResults() As Collection
   Set TestResults = m_TestResults
End Property
