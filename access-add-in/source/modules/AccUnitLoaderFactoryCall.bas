Attribute VB_Name = "AccUnitLoaderFactoryCall"
'---------------------------------------------------------------------------------------
' Modul: AccUnitLoaderFactoryCall
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Public Function GetAccUnitFactory() As AccUnitLoaderFactory
   CheckAccUnitTypeLibFile CodeVBProject
   Set GetAccUnitFactory = New AccUnitLoaderFactory
End Function
