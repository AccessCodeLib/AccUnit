VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetComDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: NetComDomain
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        .NET-COM-dll ohne Registrierung starten
' \ingroup COM
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>COM/NetComDomain.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>mscoree</name><major>2</major><minor>0</minor><guid>{5477469E-83B1-11D2-8B49-00A0C9B7C9C4}</guid></ref>
'  <ref><name>mscorlib</name><major>2</major><minor>0</minor><guid>{BED7F4EA-1A96-11D2-8F08-00A0C9A6186D}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private m_LibraryFile As String
Private m_DefaultNameSpace As String

'---------------------------------------------------------------------------------------
' Function: CreateObject
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erstellt eine neue Instanz einer Klasse aus einer .Net Com Dll
' </summary>
' <param name="ClassIdentifier">Name der Klasse ohne Namespace</param>
' <param name="DefaultNameSpace">Namespace in dem sich die Klasse befindet</param>
' <param name="LibraryFile">absoluter Pfad zur .Net Com Dll</param>
' <returns>Object</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateObject(ByVal ClassIdentifier As String, _
                    Optional ByVal DefaultNameSpace As String = vbNullString, _
                    Optional ByVal LibraryFile As String = vbNullString) As Object

   If VBA.Len(DefaultNameSpace) > 0 Then Me.DefaultNameSpace = DefaultNameSpace
   If VBA.Len(LibraryFile) > 0 Then Me.LibraryFile = LibraryFile

   Dim NetRTH As mscoree.CorRuntimeHost
   Dim NetAppDomain As mscorlib.AppDomain

   If VBA.InStr(1, ClassIdentifier, ".") = 0 Then
      If VBA.Len(Me.DefaultNameSpace) = 0 Then
         Err.Raise vbObjectError, "NetComDomain.CreateObject", "Namespace muss angegeben werden!"
      End If
      ClassIdentifier = Me.DefaultNameSpace & "." & ClassIdentifier
   End If

   Set NetRTH = New mscoree.CorRuntimeHost
   NetRTH.Start
   NetRTH.GetDefaultDomain NetAppDomain

   Set CreateObject = NetAppDomain.CreateInstanceFrom(Me.LibraryFile, ClassIdentifier).Unwrap

   NetRTH.Stop
   Set NetAppDomain = Nothing
   Set NetRTH = Nothing

End Function

Public Property Get LibraryFile() As String
    LibraryFile = m_LibraryFile
End Property

Public Property Let LibraryFile(ByVal FilePath As String)
    m_LibraryFile = FilePath
End Property

Public Property Get DefaultNameSpace() As String
    DefaultNameSpace = m_DefaultNameSpace
End Property

Public Property Let DefaultNameSpace(ByVal NameSpace As String)
    m_DefaultNameSpace = NameSpace
End Property


