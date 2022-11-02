const AddInName = "ACLib-AccUnit-Loader"
const AddInFileName = "AccUnitLoader.accda"
const MsgBoxTitle = "Update ACLib-AccUnit-Loader"

MsgBox "Vor dem Aktualisieren der Add-In-Datei darf das Add-In nicht geladen sein!" & chr(13) & _
       "Zur Sicherheit alle Access-Instanzen schlieﬂen.", , MsgBoxTitle & ": Hinweis"

Select Case MsgBox("Soll das Add-In als MDE verwendet werden?" + chr(13) & _
                   "(Add-In wird kompiliert ins Add-In-Verzeichnis kopiert.)", 3, MsgBoxTitle)
   case 6 ' vbYes
      CreateMde GetSourceFileFullName, GetDestFileFullName
   case 7 ' vbNo
	  DeleteAddInFiles
      FileCopy GetSourceFileFullName, GetDestFileFullName
   case else
      
End Select


'##################################################
' Hilfsfunktionen:

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
   Set AccessApp = CreateObject("Access.Application") 
   set fso = CreateObject("Scripting.FileSystemObject") 
   fso.CopyFile SourceFilePath, DestFilePath
End Function

Function DeleteAddInFiles()

   Set fso = CreateObject("Scripting.FileSystemObject")
   
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

Function CreateMde(SourceFilePath, DestFilePath)

   FileToCompile = DestFilePath & ".accdb"
   FileCopy SourceFilePath, FileToCompile
  
   Set AccessApp = CreateObject("Access.Application") 
   RunPrecompileProcedure AccessApp, FileToCompile
   AccessApp.SysCmd 603, (FileToCompile), (DestFilePath)
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   DeleteFile fso, FileToCompile

End Function

Function DeleteDestFiles()

	

   Set fso = CreateObject("Scripting.FileSystemObject")

   DeleteFile fso, DestFilePath
   DeleteFile fso, GetAddInLocation() & "\lib\"

End Function

Function RunPrecompileProcedure(AccessApp, SourceFilePath)

   AccessApp.OpenCurrentDatabase SourceFilePath
   AccessApp.Run "CheckAccUnitTypeLibFile"
   AccessApp.CloseCurrentDatabase

End Function
