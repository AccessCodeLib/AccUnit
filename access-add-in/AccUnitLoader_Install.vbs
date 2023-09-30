const AddInName = "ACLib-AccUnit-Loader"
const AddInFileName = "AccUnitLoader.accda"
const MsgBoxTitle = "Update ACLib-AccUnit-Loader"

MsgBox "Before updating the add-in file, the add-in must not be loaded!" & chr(13) & _
       "For safety, close all Access instances.", , MsgBoxTitle & ": Hinweis"

Dim CompletedMsg
Dim AddInFileInstalled

Select Case MsgBox("Should the add-in be used as ACCDE?" + chr(13) & _
                   "(Add-In is compiled and copied to the Add-In directory.)", 3, MsgBoxTitle)
   case 6 ' vbYes
      if CreateMde(GetSourceFileFullName, GetDestFileFullName) = True then
	      CompletedMsg = "Compiled file was created."
		  AddInFileInstalled = True
      else
          CompletedMsg = "Error! Compiled file was not created."
      end if
   case 7 ' vbNo
      DeleteAddInFiles
      if CopyFileAndRundPrecompileProc(GetSourceFileFullName, GetDestFileFullName) Then
	      CompletedMsg = "File was copied."
		  AddInFileInstalled = True
      else
	      CompletedMsg = "Error! File was not copied."
      end if
   case else
      
End Select

if AddInFileInstalled = True then
   RegisterAddIn GetDestFileFullName()
end if

MsgBox CompletedMsg, , MsgBoxTitle

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

Function CopyFileAndRundPrecompileProc(SourceFilePath, DestFilePath)

   IF Not FileCopy(SourceFilePath, DestFilePath) Then
      Exit Function
   End If

   Set AccessApp = CreateObject("Access.Application") 	
   If Not RunPrecompileProcedure(AccessApp, DestFilePath) Then
      Exit Function
   End If

   CopyFileAndRundPrecompileProc = True
 
End Function

Function CreateMde(SourceFilePath, DestFilePath)
	
   Set fso = CreateObject("Scripting.FileSystemObject")
   DeleteAddInFilesFso fso

   FileToCompile = DestFilePath & ".accdb"
   If Not FileCopy(SourceFilePath, FileToCompile) Then
      Exit Function
   End If
  
   Set AccessApp = CreateObject("Access.Application") 
   If Not RunPrecompileProcedure(AccessApp, FileToCompile) Then
      Exit Function
   End If
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
   AccessApp.Visible = True
   AccessApp.Run "CheckAccUnitTypeLibFile"
   AccessApp.CloseCurrentDatabase

   RunPrecompileProcedure = True

End Function

Function RegisterAddIn(AddInFile)

   Dim AddInDb, AccessApp, rst, ItemValue
   Dim wsh

   Set AccessApp = CreateObject("Access.Application") 
   Set AddInDb = AccessApp.DBEngine.OpenDatabase(AddInFile)
    
   set wsh = CreateObject("WScript.Shell")
   Set rst = AddInDb.OpenRecordset("select Subkey, ValName, Type, Value from USysRegInfo where ValName > '' Order By ValName", 8) 'dbOpenForwardOnly=8
   Do While Not rst.EOF
        ItemValue = rst.Fields("Value").Value
        If Len(ItemValue) > 0 Then
        If InStr(1, ItemValue, "|ACCDIR") > 0 Then
            ItemValue = AddInDb.Name
        End If
        End If
        RegisterMenuAddInItem AccessApp, wsh, rst.Fields("Subkey").Value, rst.Fields("ValName").Value, rst.Fields("Type").Value, ItemValue
        rst.MoveNext
   Loop
   rst.Close

   AddInDb.Close
	

End Function

Function RegisterMenuAddInItem(AccessApp, wsh, ByVal SubKey, ByVal ItemValName, ByVal RegType, ByVal ItemValue)
    Dim RegName
    RegName = GetRegistryPath(SubKey, AccessApp)
    With wsh
        If Len(ItemValName) > 0 Then
            RegName = RegName & "\" & ItemValName
        End If
        .RegWrite RegName, ItemValue, GetRegTypeString(RegType)
    End With
End Function

Function GetRegTypeString(ByVal RegType)
    Select Case RegType
        Case 1
            GetRegTypeString = "REG_SZ"
        Case 4
            GetRegTypeString = "REG_DWORD"
        Case 0
            GetRegTypeString = vbNullString
        Case Else
            Err.Raise vbObjectError, "GetRegTypeString", "RegType not supported"
    End Select
End Function

Function GetRegistryPath(SubKey, AccessApp)
    GetRegistryPath = Replace(SubKey, "HKEY_CURRENT_ACCESS_PROFILE", HkeyCurrentAccessProfileRegistryPath(AccessApp))
End Function

Function HkeyCurrentAccessProfileRegistryPath(AccessApp)
    HkeyCurrentAccessProfileRegistryPath = "HKCU\SOFTWARE\Microsoft\Office\" & AccessApp.Version & "\Access"
End Function


