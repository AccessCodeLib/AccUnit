using AccessCodeLib.Common.VBIDETools.Templates;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.Configuration
{
    public class AccUnitLoaderAddInCodeTemplates : CodeTemplateCollection
    {
        public AccUnitLoaderAddInCodeTemplates(bool useAccUnitTypeLib = false)
        {
            AddAccUnitLoaderFactory(useAccUnitTypeLib);
        }

        private void AddAccUnitLoaderFactory(bool useAccUnitTypeLib)
        {
            var code = AccUnitLoaderFactoryCode.Replace("{UseAccUnitTypeLib}", useAccUnitTypeLib ? "1" : "0");
            Add(new CodeTemplate(@"AccUnit_Factory", vbext_ComponentType.vbext_ct_StdModule, code));
        }

        private static readonly string AccUnitLoaderFactoryCode =
            @"Option Compare Database
Option Explicit

#Const USE_ACCUNIT_TYPELIB = {UseAccUnitTypeLib}

Private m_AccUnitLoaderFactory As Object
Private m_UseMatchResultCollector As Boolean
Private m_CodeCoverageTracker As Object

Private Function AccUnitLoaderFactory() As Object
   If m_AccUnitLoaderFactory Is Nothing Then
      Set m_AccUnitLoaderFactory = Application.Run(GetAddInPath & ""AccUnitLoader.GetAccUnitFactory"")
      If m_UseMatchResultCollector Then
         m_AccUnitLoaderFactory.Init NewDebugPrintMatchResultCollector
      End If
   End If
   Set AccUnitLoaderFactory = m_AccUnitLoaderFactory
End Function

#If USE_ACCUNIT_TYPELIB Then
Private Property Get AccUnitFactory() As AccUnit.AccUnitFactory
#Else
Private Property Get AccUnitFactory() As Object
#End If
   Set AccUnitFactory = AccUnitLoaderFactory.AccUnitFactory
End Property

Private Function GetAddInPath() As String
   GetAddInPath = Environ(""appdata"") & ""\Microsoft\AddIns\""
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

Public Function NewDebugPrintMatchResultCollector(Optional ByVal ShowPassedText As Boolean = False, Optional ByVal UseRaiseErrorForFailedMatch As Boolean = True) As Object
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
   If Not m_UseMatchResultCollector Then
      m_UseMatchResultCollector = True
      Set m_AccUnitLoaderFactory = Nothing
   End If
   Set TestRunner = AccUnitLoaderFactory.TestRunner
End Property

#If USE_ACCUNIT_TYPELIB Then
Public Property Get TestSuite() As AccUnit.VBATestSuite
#Else
Public Property Get TestSuite() As Object
#End If
   If m_UseMatchResultCollector Then
      m_UseMatchResultCollector = False
      Set m_AccUnitLoaderFactory = Nothing
   End If
   Set TestSuite = AccUnitLoaderFactory.DebugPrintTestSuite
End Property

Public Sub RunTest(ByVal testClassInstance As Object, Optional ByVal MethodName As String = ""*"", Optional ByVal PrintSummary As Boolean = True, Optional ByVal TestResultCollector As Object)
   If Not m_UseMatchResultCollector Then
      m_UseMatchResultCollector = True
      Set m_AccUnitLoaderFactory = Nothing
   End If
   AccUnitLoaderFactory.RunTest testClassInstance, MethodName, PrintSummary, TestResultCollector
End Sub

Public Sub RunAllTests()
   TestSuite.AddFromVBProject.Run
End Sub

#If USE_ACCUNIT_TYPELIB Then
Public Property Get CodeCoverageTracker(Optional ReInit As Boolean = False) As AccUnit.CodeCoverageTracker
#Else
Public Property Get CodeCoverageTracker(Optional ReInit As Boolean = False) As Object
#End If
   If ReInit Then
      If Not m_CodeCoverageTracker Is Nothing Then
         m_CodeCoverageTracker.Dispose
         Set m_CodeCoverageTracker = Nothing
      End If
   End If
   If m_CodeCoverageTracker Is Nothing Then
      Set m_CodeCoverageTracker = AccUnitLoaderFactory.CodeCoverageTracker
   End If
   Set CodeCoverageTracker = m_CodeCoverageTracker
End Property

#If USE_ACCUNIT_TYPELIB Then
Public Function CodeCoverageTest(ParamArray CodeModulNames() As Variant) As AccUnit.VBATestSuite
#Else
Public Function CodeCoverageTest(ParamArray CodeModulNames() As Variant) As Object
#End If
   Dim CodeModuleName As Variant
   Dim CodeCoverageTestSuite As Object

   With CodeCoverageTracker(True)
      For Each CodeModuleName In CodeModulNames
         .Add CodeModuleName
      Next
   End With
   
   If m_UseMatchResultCollector Then
      m_UseMatchResultCollector = False
      Set m_AccUnitLoaderFactory = Nothing
   End If
   Set CodeCoverageTestSuite = AccUnitLoaderFactory.DebugPrintTestSuite
   Set CodeCoverageTestSuite.CodeCoverageTracker = m_CodeCoverageTracker
   
   Set CodeCoverageTest = CodeCoverageTestSuite
   
End Function
";
    }
}
