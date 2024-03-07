const AddInName = "ACLib-AccUnit-Loader"
const AddInFileName = "AccUnitLoader.xlam"
const MsgBoxTitle = "Install ACLib-AccUnit-Loader"

Dim AddInFileInstalled, CompletedMsg

MsgBox "Before updating the add-in file, the add-in must not be loaded!" & chr(13) & _
       "For safety, close all Excel instances.", , MsgBoxTitle & ": Information"

DeleteAddInFiles
AddInFileInstalled = CopyFiles()
If AddInFileInstalled Then
  CompletedMsg = "Add-In was saved in '" + GetAddInLocation + "'." & chr(13) & _
		 "Next step: open Excel and activate Add-In."
Else
  CompletedMsg = "Error! File was not copied."
End If

If AddInFileInstalled = True Then
'   RegisterAddIn GetDestFileFullName()
End If

If CompletedMsg > "" Then
   MsgBox CompletedMsg, , MsgBoxTitle
End If


'##################################################
' Functions

Function GetSourceFileFullName()
   GetSourceFileFullName = GetScriptLocation & AddInFileName 
End Function

Function GetDestFileFullName()
   GetDestFileFullName = GetAddInLocation & AddInFileName 
End Function

Function GetScriptLocation()
   With WScript
      GetScriptLocation = Replace(.ScriptFullName & ":", .ScriptName & ":", "") 
   End With
End Function

Function GetAddInLocation()
   GetAddInLocation = GetAppDataLocation & "Microsoft\AddIns\"
End Function

Function GetAppDataLocation()
   Set wsShell = CreateObject("WScript.Shell")
   GetAppDataLocation = wsShell.ExpandEnvironmentStrings("%APPDATA%") & "\"
End Function

Function FileCopy(SourceFilePath, DestFilePath)
   set fso = CreateObject("Scripting.FileSystemObject") 
   fso.CopyFile SourceFilePath, DestFilePath
   FileCopy = True
End Function

Function DeleteAddInFiles()
   Set fso = CreateObject("Scripting.FileSystemObject")
   DeleteAddInFilesFso fso
End Function

Function DeleteAddInFilesFso(fso)

   DestFile = GetDestFileFullName()
   Tlbfile = GetAddInLocation & "\lib\AccessCodeLib.AccUnit.tlb"
   
   DeleteFile fso, DestFile
   DeleteFile fso, Tlbfile
   
End Function

Function DeleteFile(fso, File2Delete)
   if fso.FileExists(File2Delete) then
      fso.DeleteFile File2Delete
   end if
End Function

Function CopyFiles()

   'Add-In:
   IF Not FileCopy(GetSourceFileFullName, GetDestFileFullName) Then
      Exit Function
   End If

   CopyFiles = True
 
End Function

'Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Add-in Manager
