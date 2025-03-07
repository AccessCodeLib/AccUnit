﻿Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
' Modul: _config_Application
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'<codelib>
'  <license>_codelib/license.bas</license>
'  <use>%AppFolder%/source/defGlobal_AccUnitLoader.bas</use>
'  <use>base/modApplication.bas</use
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/ApplicationHandler_AppFile.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>_codelib/addins/shared/AccUnitConfiguration.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'Version number
Private Const APPLICATION_VERSION As String = "0.9.1101.250216"

Private Const APPLICATION_NAME As String = "ACLib AccUnit Loader"
Private Const APPLICATION_FULLNAME As String = "Access Code Library - AccUnit Loader"
Private Const APPLICATION_TITLE As String = APPLICATION_FULLNAME
Private Const APPLICATION_ICONFILE As String = "ACLib.ico"
Public Const ACCUNIT_TYPELIB_FILE As String = "AccUnit.tlb"
Public Const ACCUNIT_DLL_FILE As String = "AccUnit.dll"

Private Const APPLICATION_STARTFORMNAME As String = "AccUnitLoaderForm"

Private m_Extensions As Object 'ApplicationHandler_ExtensionCollection

'---------------------------------------------------------------------------------------
' Sub: InitConfig
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Konfigurationseinstellungen initialisieren
' </summary>
' <param name="oCurrentAppHandler">Möglichkeit einer Referenzübergabe, damit nicht CurrentApplication genutzt werden muss</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitConfig(Optional ByRef CurrentAppHandlerRef As Object = Nothing)

'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   defGlobal_AccUnitLoader.ACLibIconFileName = APPLICATION_ICONFILE

'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If CurrentAppHandlerRef Is Nothing Then
      Set CurrentAppHandlerRef = modApplication.CurrentApplication
   End If

   With CurrentAppHandlerRef

      'Zur Sicherheit AccDb einstellen
      Set .AppDb = Application.CodeDb 'muss auf CodeDb zeigen,
                          'da diese Anwendung als Add-In verwendet wird

      'Anwendungsname
      .ApplicationName = APPLICATION_NAME
      .ApplicationFullName = APPLICATION_FULLNAME
      .ApplicationTitle = APPLICATION_TITLE

      'Version
      .Version = APPLICATION_VERSION

      ' Formular, das am Ende von CurrentApplication.Start aufgerufen wird
      .ApplicationStartFormName = APPLICATION_STARTFORMNAME

   End With

'----------------------------------------------------------------------------
' Erweiterungen:
'
   Set m_Extensions = New ApplicationHandler_ExtensionCollection
   With m_Extensions
      Set .ApplicationHandler = CurrentAppHandlerRef
      .Add New ApplicationHandler_AppFile
   End With

End Sub


'############################################################################
'
' Funktionen für die Anwendungswartung
' (werden nur im Anwendungsentwurf benötigt)
'
'----------------------------------------------------------------------------
' Hilfsfunktion zum Speichern von Dateien in die lokale AppFile-Tabelle
'----------------------------------------------------------------------------
'Private Sub SetAppFiles()
'
'   Dim accFileName As Variant
'
'  ' Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)
'   With modApplication.CurrentApplication.Extensions("AppFile")
'      For Each accFileName In AccUnitLoaderConfigProcedures.AccUnitFileNames
'         .SaveAppFile accFileName, CodeProject.Path & "\lib\" & accFileName, True
'      Next
'   End With
'
'End Sub

Public Sub PrepareForVCS()
   If DaoTools.TableDefExists("ACLib_ConfigTable") Then
      Application.CurrentDb.TableDefs.Delete "ACLib_ConfigTable"
      Application.RefreshDatabaseWindow
   End If
   AccUnitLoaderConfigProcedures.RemoveAccUnitTlbReference
End Sub
