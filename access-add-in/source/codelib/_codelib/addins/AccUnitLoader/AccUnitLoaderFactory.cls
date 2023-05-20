VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccUnitLoaderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Klassenmodul: AccUnitLoaderFactory
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Factory-Klasse für AccUnit
' </summary>
' <remarks>
' </remarks>
'\ingroup addins_shared
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/AccUnitLoader/AccUnitLoaderFactory.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---
Option Compare Database
Option Explicit

Private m_AccUnitFactory As Object
Private m_MatchResultCollector As Object

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   If Not (m_MatchResultCollector Is Nothing) Then
      Set m_MatchResultCollector = Nothing
   End If
   
   If Not (m_AccUnitFactory Is Nothing) Then
      'm_AccUnitFactory.Dispose ' COM-sichtbar machen?
      Set m_AccUnitFactory = Nothing
   End If

End Sub

Public Property Get AccUnitFactory() As Object 'AccUnit.AccUnitFactory
   If m_AccUnitFactory Is Nothing Then
      InitAccUnitFactory
   End If
   Set AccUnitFactory = m_AccUnitFactory
End Property

Private Sub InitAccUnitFactory()

   With New NetComDomain
      .DefaultNameSpace = "AccessCodeLib.AccUnit.Interop"
      .LibraryFile = CurrentAccUnitConfiguration.AccUnitDllPath & ACCUNIT_DLL_FILE
      Set m_AccUnitFactory = .CreateObject("AccUnitFactory")
   End With

End Sub

Public Sub Init(ByVal AssertMatchResultCollector As Object)
   Set m_MatchResultCollector = AssertMatchResultCollector
End Sub

Public Function Assert() As Object ' AccUnit.Assert

   Set Assert = AccUnitFactory.Assert
   If Not (m_MatchResultCollector Is Nothing) Then
      Set Assert.MatchResultCollector = m_MatchResultCollector
   End If
   
End Function

Public Property Get ConstraintBuilder() As Object ' AccUnit.ConstraintBuilder
    Set ConstraintBuilder = AccUnitFactory.ConstraintBuilder
End Property

Public Function NewDebugPrintMatchResultCollector(Optional ByVal ShowPassedText As Boolean = False, Optional ByVal UseRaiseErrorForFailedMatch As Boolean = True) As DebugPrintMatchResultCollector

   With New DebugPrintMatchResultCollector
      .Init ShowPassedText, UseRaiseErrorForFailedMatch
      Set NewDebugPrintMatchResultCollector = .This
   End With
   
End Function

Public Function NewDebugPrintTestResultCollector() As DebugPrintTestResultCollector
   Set NewDebugPrintTestResultCollector = New DebugPrintTestResultCollector
End Function

Public Property Get TestBuilder() As Object 'AccUnit.TestBuilder
    Set TestBuilder = AccUnitFactory.TestBuilder
    Set TestBuilder.HostApplication = Application
End Property

Public Property Get TestRunner() As Object 'AccUnit.TestRunner
    
    SetErrorTrappingToBreakOnUnhandledErrors
    Set TestRunner = AccUnitFactory.TestRunner(CurrentVbProject)
    
End Property

Private Sub SetErrorTrappingToBreakOnUnhandledErrors()
'BreakOnAllErrors = 0,
'BreakInClassModule = 1,
'BreakOnUnhandledErrors = 2
   Const ErrorTrappingOptionName As String = "Error Trapping"
   
   If Application.GetOption(ErrorTrappingOptionName) <> 2 Then
      Application.SetOption ErrorTrappingOptionName, 2
      Debug.Print "Error Trapping auf BreakOnUnhandledErrors (2) umgestellt"
   End If
            
End Sub

Public Sub RunTest(ByVal testClassInstance As Object, Optional ByVal MethodName As String = "*", Optional ByVal PrintSummary As Boolean = True, Optional ByVal TestResultCollector As Object)

   If PrintSummary And TestResultCollector Is Nothing Then
      Set TestResultCollector = New DebugPrintTestResultCollector
   End If
   
   TestRunner.Run testClassInstance, MethodName, TestResultCollector
   
   If PrintSummary Then
      TestResultCollector.PrintSummary True
   End If
   
End Sub

Public Property Get Configurator() As Object 'AccUnit.Configurator
   Set Configurator = AccUnitFactory.Configurator
End Property

Public Property Get VbaTestSuite() As AccUnit.VbaTestSuite
   Set VbaTestSuite = AccUnitFactory.VbaTestSuite
   
   With VbaTestSuite
      Set .HostApplication = Application
      Set .ActiveVBProject = CurrentVbProject
      Set .TestRunner = TestRunner
   End With
   
End Property

Public Property Get DebugPrintTestSuite() As DebugPrintTestSuite

   Set DebugPrintTestSuite = New DebugPrintTestSuite
   
   With DebugPrintTestSuite
      .Init VbaTestSuite, NewDebugPrintTestResultCollector
   End With
   
End Property

Public Property Get CodeCoverageTracker() As AccUnit.CodeCoverageTracker
    Set CodeCoverageTracker = AccUnitFactory.CodeCoverageTracker(CurrentVbProject)
End Property
