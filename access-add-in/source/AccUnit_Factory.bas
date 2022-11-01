Attribute VB_Name = "AccUnit_Factory"
Option Compare Database
Option Explicit

#Const USE_ACCUNIT_TYPELIB = 0

Private m_AccUnitLoaderFactory As Object

Private Property Get AccUnitLoaderFactory() As Object
   If m_AccUnitLoaderFactory Is Nothing Then
      Set m_AccUnitLoaderFactory = Application.Run(GetAddInPath & "AccUnitLoader.GetAccUnitFactory")
   End If
   Set AccUnitLoaderFactory = m_AccUnitLoaderFactory
End Property

Private Function GetAddInPath() As String
   GetAddInPath = Environ("appdata") & "\Microsoft\AddIns\"
End Function

#If USE_ACCUNIT_TYPELIB Then
Public Property Get Assert() As AccUnit.Assert
#Else
Public Property Get Assert() As Object
#End If
   Set Assert = AccUnitLoaderFactory.Assert
End Property

#If USE_ACCUNIT_TYPELIB Then
Public Property Get Iz() As AccUnit.ConstraintBuilder
#Else
Public Property Get Iz() As Object
#End If
    Set Iz = AccUnitLoaderFactory.ConstraintBuilder
End Property

#If USE_ACCUNIT_TYPELIB Then
Public Property Get TestBuilder() As AccUnit.TestBuilder
#Else
Public Property Get TestBuilder() As Object
#End If
    Set TestBuilder = AccUnitLoaderFactory.TestBuilder
End Property

Public Function NewDebugPrintMatchResultCollector(Optional ByVal ShowPassedText As Boolean = False, Optional ByVal UseRaiseErrorForFailedMatch As Boolean = False) As Object
   Set NewDebugPrintMatchResultCollector = AccUnitLoaderFactory.NewDebugPrintMatchResultCollector(ShowPassedText, UseRaiseErrorForFailedMatch)
End Function

Public Function NewDebugPrintTestResultCollector() As Object
   Set NewDebugPrintTestResultCollector = AccUnitLoaderFactory.NewDebugPrintTestResultCollector
End Function

#If USE_ACCUNIT_TYPELIB Then
Public Property Get TestRunner() As AccUnit.TestRunner
#Else
Public Property Get TestRunner() As Object
#End If
   Set TestRunner = AccUnitLoaderFactory.TestRunner
End Property

Public Sub RunTest(ByVal testClassInstance As Object, Optional ByVal MethodName As String = "*", Optional ByVal PrintSummary As Boolean = True, Optional ByVal TestResultCollector As Object)
   AccUnitLoaderFactory.RunTest testClassInstance, MethodName, PrintSummary, TestResultCollector
End Sub

