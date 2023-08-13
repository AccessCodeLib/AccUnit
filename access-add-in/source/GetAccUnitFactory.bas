Attribute VB_Name = "AccUnitLoaderFactoryCall"
'---------------------------------------------------------------------------------------
' Modul: AccUnitLoaderFactoryCall
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/GetAccUnitFactory.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>%AppFolder%/source/AccUnitLoaderFactory.cls</use>
'</codelib>
'---
Option Compare Database
Option Explicit

Public Function GetAccUnitFactory() As AccUnitLoaderFactory
   CheckAccUnitTypeLibFile
   Set GetAccUnitFactory = New AccUnitLoaderFactory
End Function
