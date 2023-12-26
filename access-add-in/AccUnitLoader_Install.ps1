$AddInName = "ACLib-AccUnit-Loader"
$AddInFileName = "AccUnitLoader.accda"
$MsgBoxTitle = "Update ACLib-AccUnit-Loader"


function Get-SourceFileFullName {
   $ScriptLocation = Get-ScriptLocation 
   $retVal = '' + $ScriptLocation + $AddInFileName
   $retVal
}

function Get-DestFileFullName {
   $AddInLocation = Get-AddInLocation
   $retVal = $AddInLocation + $AddInFileName
   $retVal
}

function Get-ScriptLocation {
   $retVal = '' + $PSScriptRoot + "\"
   $retVal
}

function Get-AddInLocation {
   $env:APPDATA + "\Microsoft\AddIns\"
}

function FileCopy {
   Param($SourceFilePath, $DestFilePath)
   Copy-Item $SourceFilePath -Destination $DestFilePath
   $true
}

function Delete-AddInFiles {

  $DestFile = Get-DestFileFullName
  $AddInLocation = Get-AddInLocation 
  $tlbPath = "\lib\AccessCodeLib.AccUnit.tlb"
  $Tlbfile = $AddInLocation + $tlbPath
   
 # Remove-Item -Path $DestFile -Force 
 # Remove-Item -Path $Tlbfile -Force 
   
}

function CopyFileAndRunPrecompileProc {

    Param($SourceFilePath, $DestFilePath)

    [System.Windows.Forms.MessageBox]::Show($SourceFilePath, "SourceFilePath", 0)
    

    #Copy-Item -Path $SourceFilePath -Destination $DestFilePath

    $ret = Run-PrecompileProcedure $DestFilePath
    If ($ret -eq $true) {
        $true
    }
    else {
        $false
    }
}

function CreateMde {
    Param($SourceFilePath, $DestFilePath)
	
   Delete-AddInFiles

   $FileToCompile = $DestFilePath + ".accdb"
   $copied = FileCopy $SourceFilePath $FileToCompile
   If ($copied -ne $true) {
      return
   }
  
   #$AccessApp = New-Object -ComObject Access.Application 
   $ret = Run-PrecompileProcedure $FileToCompile
   If ($ret -ne $true) {
      return
   }

   return

   $AccessApp.SysCmd(603, ($FileToCompile), ($DestFilePath))
   
   Remove-Item -Path $FileToCompile -Force 

   $true

}

function Run-PrecompileProcedure {
    Param([string]$AccdbFilePath)

   $AccessApp = New-Object -ComObject Access.Application
   $AccessApp.Visible = $true

   $AccessApp.OpenCurrentDatabase("" + $AccdbFilePath)
   
   $RunMethod = "CheckAccUnitTypeLibFile"
   $AccessApp.Run($RunMethod)
   $AccessApp.CloseCurrentDatabase

   $true

}


####main


$msg = "Before updating the add-in file, the add-in must not be loaded!
For safety, close all Access instances."

[System.Windows.Forms.MessageBox]::Show($msg, $MsgBoxTitle + ": Information", 0)

$msg ="Should the add-in be used as a compiled file (accde)?
(Add-In is compiled and copied to the Add-In directory.)"

$msgRet = [System.Windows.Forms.MessageBox]::Show($msg, $MsgBoxTitle, 3)

$AddInLocation = Get-AddInLocation

switch ( $msgRet )
{
    "Yes" { 
        $AddInFileInstalled = $true # CreateMde(GetSourceFileFullName, GetDestFileFullName)

        If ( $AddInFileInstalled -eq $true ) {
	      $CompletedMsg = "Add-In was compiled and saved in '" + $AddInLocation + "'."
        }
        Else {
          $CompletedMsg = "Error! Compiled file was not created."
        } 
    
    }

    "No" { 
        #Delete-AddInFiles
        $SourceFileFullName = Get-SourceFileFullName
        $DestFileFullName = Get-DestFileFullName
	    $AddInFileInstalled = CopyFileAndRunPrecompileProc $SourceFileFullName $DestFileFullName
        If ( $AddInFileInstalled ) {
	      $CompletedMsg = "Add-In was saved in '" + $AddInLocation + "'."
        }
        Else {
	      $CompletedMsg = "Error! File was not copied."
        }
    }

    default  { 
        return
    }
}

[System.Windows.Forms.MessageBox]::Show($CompletedMsg, $MsgBoxTitle, 0)

