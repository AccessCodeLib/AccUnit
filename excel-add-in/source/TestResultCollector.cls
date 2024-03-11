VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResultCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: TestResultCollector
'---------------------------------------------------------------------------------------
'
'
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Implements AccUnit.ITestResultCollector
Implements AccUnit.ITestSummaryPrinter

Private m_TestResults As Collection
Private m_TestSummary As TestSummary

Public Event TestSuiteStarted(ByVal TestSuite As AccUnit.ITestSuite)
Public Event TestTraceMessage(ByVal Message As String, ByVal CodeCoverageTracker As CodeCoverageTracker)
Public Event NewTestResult(ByVal TestResult As AccUnit.ITestResult)
Public Event TestSuiteFinished(ByVal TestSummary As AccUnit.ITestSummary)
Public Event PrintSummary(ByVal TestSummary As AccUnit.ITestSummary, ByVal PrintTestResults As Boolean)

Private Sub Class_Initialize()
   Set m_TestResults = New Collection
   Set m_TestSummary = New TestSummary
End Sub

Public Sub TestSuiteStarted(ByVal TestSuite As AccUnit.ITestSuite)
    RaiseEvent TestSuiteStarted(TestSuite)
End Sub

Public Sub PrintTestTraceMessage(ByVal Message As String, ByVal CodeCoverageTracker As CodeCoverageTracker)
   RaiseEvent TestTraceMessage(Message, CodeCoverageTracker)
End Sub

Private Sub ITestResultCollector_Add(ByVal TestResult As AccUnit.ITestResult)
   m_TestResults.Add TestResult
   m_TestSummary.SaveResult TestResult
   RaiseEvent NewTestResult(TestResult)
End Sub

Public Sub TestSuiteFinished(ByVal TestSummary As AccUnit.ITestSummary)
    RaiseEvent TestSuiteFinished(TestSummary)
End Sub

Private Property Get ITestSummaryPrinter_Name() As String
   ITestSummaryPrinter_Name = "TestResultCollectorEventBridge"
End Property

Private Sub ITestSummaryPrinter_PrintSummary(Optional ByVal PrintTestResults As Boolean = False)
   RaiseEvent PrintSummary(m_TestSummary, PrintTestResults)
End Sub

Public Sub PrintSummary(ByVal TestSummary As AccUnit.ITestSummary, Optional ByVal PrintTestResults As Boolean = False)
   RaiseEvent PrintSummary(TestSummary, PrintTestResults)
End Sub

Public Property Get Summary() As TestSummary
   Set Summary = m_TestSummary
End Property

Public Property Get TestResults() As Collection
   Set TestResults = m_TestResults
End Property


