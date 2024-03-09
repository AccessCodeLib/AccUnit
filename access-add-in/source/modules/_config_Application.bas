Attribute VB_Name = "_config_Application"
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
Option Compare Database
Option Explicit

'Version nummer
Private Const APPLICATION_VERSION As String = "0.9.21.240308"

Private Const APPLICATION_NAME As String = "ACLib AccUnit Loader"
Private Const APPLICATION_FULLNAME As String = "Access Code Library - AccUnit Loader"
Private Const APPLICATION_TITLE As String = APPLICATION_FULLNAME
Private Const APPLICATION_ICONFILE As String = "ACLib.ico"
Public Const ACCUNIT_TYPELIB_FILE As String = "AccessCodeLib.AccUnit.tlb"
Public Const ACCUNIT_DLL_FILE As String = "AccessCodeLib.AccUnit.dll"

Private Const APPLICATION_STARTFORMNAME As String = "AccUnitLoaderForm"

Private m_Extensions As ApplicationHandler_ExtensionCollection

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
Public Sub InitConfig(Optional ByRef CurrentAppHandlerRef As ApplicationHandler = Nothing)

'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   defGlobal_AccUnitLoader.ACLibIconFileName = APPLICATION_ICONFILE

'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If CurrentAppHandlerRef Is Nothing Then
      Set CurrentAppHandlerRef = CurrentApplication
   End If

   With CurrentAppHandlerRef

      'Zur Sicherheit AccDb einstellen
      Set .AppDb = CodeDb 'muss auf CodeDb zeigen,
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
      .Add New AccUnitConfiguration
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
Private Sub SetAppFiles()

   Dim accFileName As Variant

  ' Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)
   With CurrentApplication.Extensions("AppFile")
      For Each accFileName In AccUnitFileNames
         .SaveAppFile accFileName, CodeProject.Path & "\lib\" & accFileName, True
      Next
   End With



End Sub
