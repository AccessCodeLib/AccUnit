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
   Debug.Print , "Ignored: "; m_TestSummary.Ignored
   Debug.Print String(40, "-")
End Sub

Private Sub PrintTestDetailSummary()
   
   Dim i As Long
   Dim r As AccUnit.ITestResult
   For Each r In m_TestResults
      i = i + 1
      Debug.Print i, r.test.FullName, r.Result, r.Message
   Next

End Sub

Public Property Get Summary() As TestSummary
   Set Summary = m_TestSummary
End Property

Public Property Get TestResults() As Collection
   Set TestResults = m_TestResults
End Property
