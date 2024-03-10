﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Class: TestSuite
'---------------------------------------------------------------------------------------
'
' AccUnitLoader TestSuite for Access
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Implements AccUnit.IAccessTestSuite

Private WithEvents m_TestSuite As AccUnit.AccessTestSuite
Attribute m_TestSuite.VB_VarHelpID = -1
Private m_TestResultCollector As TestResultCollector
Private m_CodeCoverageTracker As AccUnit.CodeCoverageTracker

Private m_TestResultReporters As Collection

Private Sub Class_Initialize()
   Set m_TestResultReporters = New Collection
End Sub

Private Sub Class_Terminate()
   m_TestSuite.Dispose
   Set m_TestSuite = Nothing
End Sub

Public Sub Init(ByVal NewTestSuite As AccUnit.AccessTestSuite, _
                ByVal NewTestResultCollector As TestResultCollector)
   Set m_TestSuite = NewTestSuite
   Set m_TestResultCollector = NewTestResultCollector
   Set m_TestSuite.TestResultCollector = m_TestResultCollector
End Sub

Public Sub AddTestResultReporter(ByVal Reporter As ITestResultReporter)
   Set Reporter.TestResultCollector = m_TestResultCollector
   m_TestResultReporters.Add Reporter
End Sub

Public Property Get TestSuite() As AccUnit.AccessTestSuite
   Set TestSuite = m_TestSuite
End Property

Private Sub m_TestSuite_TestTraceMessage(ByVal Message As String)
   m_TestResultCollector.PrintTestTraceMessage Message, m_CodeCoverageTracker
End Sub

Public Sub PrintSummary(Optional ByVal PrintTestResults As Boolean = False)
   m_TestResultCollector.PrintSummary PrintTestResults
End Sub

Public Property Set CodeCoverageTracker(ByVal NewRef As Object)
   Set m_CodeCoverageTracker = NewRef
End Property

'####################################
' IAccessTestSuite
'
Private Function IAccessTestSuite_Reset(Optional ByVal mode As AccUnit.ResetMode = 1&) As AccUnit.IAccessTestSuite
   m_TestSuite.Reset mode
   Set IAccessTestSuite_Reset = Me
End Function

Private Property Set IAccessTestSuite_ActiveVBProject(ByVal RHS As VBIDE.VBProject)
   Set m_TestSuite.ActiveVBProject = RHS
End Property

Private Property Get IAccessTestSuite_ActiveVBProject() As VBIDE.VBProject
   Set IAccessTestSuite_ActiveVBProject = m_TestSuite.ActiveVBProject
End Property

Public Function Add(ByVal testToAdd As Object) As TestSuite
   m_TestSuite.Add testToAdd
   Set Add = Me
End Function

Private Function IAccessTestSuite_Add(ByVal testToAdd As Object) As AccUnit.IAccessTestSuite
   Set IAccessTestSuite_Add = Add(testToAdd)
End Function

Public Function AddByClassName(ByVal ClassName As String) As TestSuite
   m_TestSuite.AddByClassName ClassName
   Set AddByClassName = Me
End Function

Private Function IAccessTestSuite_AddByClassName(ByVal ClassName As String) As AccUnit.IAccessTestSuite
   Set IAccessTestSuite_AddByClassName = AddByClassName(ClassName)
End Function

Public Function AddFromVBProject() As TestSuite
   m_TestSuite.AddFromVBProject
   Set AddFromVBProject = Me
End Function

Private Function IAccessTestSuite_AddFromVBProject() As AccUnit.IAccessTestSuite
   Set IAccessTestSuite_AddFromVBProject = AddFromVBProject()
End Function

Private Sub IAccessTestSuite_Dispose()
   m_TestSuite.Dispose
End Sub

Private Property Set IAccessTestSuite_HostApplication(ByVal RHS As Object)
   Set m_TestSuite.HostApplication = RHS
End Property

Private Property Get IAccessTestSuite_HostApplication() As Object
   Set IAccessTestSuite_HostApplication = m_TestSuite.HostApplication
End Property

Private Property Get IAccessTestSuite_Name() As String
   IAccessTestSuite_Name = "TestSuite"
End Property

Public Function Run() As TestSuite
   m_TestSuite.Run
   Set Run = Me
End Function

Private Function IAccessTestSuite_Run() As AccUnit.IAccessTestSuite
   Set IAccessTestSuite_Run = Run()
End Function

Public Function SelectTests(ByVal TestNameFilter As Variant) As TestSuite
   m_TestSuite.SelectTests TestNameFilter
   Set SelectTests = Me
End Function

Private Function IAccessTestSuite_SelectTests(ByVal TestNameFilter As Variant) As AccUnit.IAccessTestSuite
   Set IAccessTestSuite_SelectTests = SelectTests(TestNameFilter)
End Function

Public Function Filter(ByVal FilterTags As Variant) As TestSuite
   m_TestSuite.Filter FilterTags
   Set Filter = Me
End Function

Private Function IAccessTestSuite_Filter(ByVal FilterTags As Variant) As AccUnit.IAccessTestSuite
   Set IAccessTestSuite_Filter = Filter(FilterTags)
End Function

Private Property Get IAccessTestSuite_Summary() As AccUnit.ITestSummary
   Set IAccessTestSuite_Summary = m_TestSuite.Summary
End Property

Private Property Set IAccessTestSuite_TestResultCollector(ByVal RHS As AccUnit.ITestResultCollector)
   Set m_TestSuite.TestResultCollector = RHS
End Property

Private Property Get IAccessTestSuite_TestResultCollector() As AccUnit.ITestResultCollector
   Set IAccessTestSuite_TestResultCollector = m_TestSuite.TestResultCollector
End Property

Private Property Set IAccessTestSuite_TestRunner(ByVal RHS As AccUnit.ITestRunner)
   Set m_TestSuite.TestRunner = RHS
End Property

Private Property Get IAccessTestSuite_TestRunner() As AccUnit.ITestRunner
   Set IAccessTestSuite_TestRunner = m_TestSuite.TestRunner
End Property

Private Property Get IAccessTestSuite_TestClassGenerator() As AccUnit.ITestClassGenerator
   Set IAccessTestSuite_TestClassGenerator = m_TestSuite.TestClassGenerator
End Property