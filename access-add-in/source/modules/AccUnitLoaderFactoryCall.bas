Attribute VB_Name = "AccUnitLoaderFactoryCall"
'---------------------------------------------------------------------------------------
' Modul: AccUnitLoaderFactoryCall
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Public Function GetAccUnitFactory() As AccUnitLoaderFactory
   modTypeLibCheck.CheckAccUnitTypeLibFile modVbProject.CodeVBProject
   Set GetAccUnitFactory = New AccUnitLoaderFactory
End Function
