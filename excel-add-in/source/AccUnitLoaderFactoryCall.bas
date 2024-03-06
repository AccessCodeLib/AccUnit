Attribute VB_Name = "AccUnitLoaderFactoryCall"
'---------------------------------------------------------------------------------------
' Modul: AccUnitLoaderFactoryCall
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/AccUnitLoaderFactoryCall.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>%AppFolder%/source/AccUnitLoaderFactory.cls</use>
'</codelib>
'---
Option Compare Text
Option Explicit

Public Function GetAccUnitFactory() As AccUnitLoaderFactory
   CheckAccUnitTypeLibFile CodeVBProject
   Set GetAccUnitFactory = New AccUnitLoaderFactory
End Function

