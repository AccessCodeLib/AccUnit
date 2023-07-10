const AddInName = "ACLib-AccUnit-Loader"
const AddInFileName = "AccUnitLoader.accda"
const MsgBoxTitle = "Update ACLib-AccUnit-Loader"

MsgBox "Before updating the add-in file, the add-in must not be loaded!" & chr(13) & _
       "For safety, close all Access instances.", , MsgBoxTitle & ": Hinweis"

Select Case MsgBox("Should the add-in be used as ACCDE?" + chr(13) & _
                   "(Add-In is compiled and copied to the Add-In directory.)", 3, MsgBoxTitle)
   case 6 ' vbYes
      if CreateMde(GetSourceFileFullName, GetDestFileFullName) = True then
	  MsgBox "Compiled file was created.", , MsgBoxTitle
      else
          MsgBox "Error! Compiled file was not created.", , MsgBoxTitle
      end if
   case 7 ' vbNo
      DeleteAddInFiles
      if FileCopy(GetSourceFileFullName, GetDestFileFullName) Then
	  MsgBox "File was copied.", , MsgBoxTitle
      else
	  MsgBox "Error! File was not copied.", , MsgBoxTitle
      end if
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

Function CreateMde(SourceFilePath, DestFilePath)
	
   Set fso = CreateObject("Scripting.FileSystemObject")
   DeleteAddInFilesFso fso

   FileToCompile = DestFilePath & ".accdb"
   FileCopy SourceFilePath, FileToCompile
  
   Set AccessApp = CreateObject("Access.Application") 
   RunPrecompileProcedure AccessApp, FileToCompile
   AccessApp.SysCmd 603, (FileToCompile), (DestFilePath)
   
   DeleteFile fso, FileToCompile

   CreateMde = True

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
