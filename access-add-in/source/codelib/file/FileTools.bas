Attribute VB_Name = "FileTools"
Attribute VB_Description = "Funktionen für Dateioperationen"
'---------------------------------------------------------------------------------------
' Module: FileTools
'---------------------------------------------------------------------------------------
'/**
'\author    Josef Poetzl
'\short     Funktionen für Dateioperationen
' <remarks>
' </remarks>
'\ingroup file
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>file/FileTools.bas</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/file/FileToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

Private Const SELECTBOX_FILE_DIALOG_TITLE As String = "Datei auswählen"
Private Const SELECTBOX_FOLDER_DIALOG_TITLE As String = "Ordner auswählen"
Private Const SELECTBOX_OPENTITLE As String = "auswählen"

Private Const DEFAULT_TEMPPATH_NOENV As String = "C:\"
Private Const PATHLEN_MAX As Long = 255

Private Const SE_ERR_NOTFOUND As Long = 2
Private Const SE_ERR_NOASSOC  As Long = 31

#If VBA7 Then

Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private Declare PtrSafe Function API_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Declare PtrSafe Function API_GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" ( _
         ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long
         
Private Declare PtrSafe Function API_ShellExecuteA Lib "shell32.dll" ( _
         ByVal Hwnd As LongPtr, _
         ByVal lOperation As String, _
         ByVal lpFile As String, _
         ByVal lpParameters As String, _
         ByVal lpDirectory As String, _
         ByVal nShowCmd As Long) As Long

#Else

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private Declare Function API_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Declare Function API_GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" ( _
         ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long

Private Declare Function API_ShellExecuteA Lib "shell32.dll" ( _
         ByVal Hwnd As Long, _
         ByVal lOperation As String, _
         ByVal lpFile As String, _
         ByVal lpParameters As String, _
         ByVal lpDirectory As String, _
         ByVal nShowCmd As Long) As Long

#End If

'---------------------------------------------------------------------------------------
' Function: SelectFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei mittels Dialog auswählen
' </summary>
' <param name="InitDir">Startverzeichnis</param>
' <param name="DlgTitle">Dialogtitel</param>
' <param name="FilterString">Filterwerten - Beispiel: "(*.*)" oder "Alle (*.*)|Textdateien (*.txt)|Bilder (*.png;*.jpg;*.gif)</param>
' <param name="MultiSelect">Mehrfachauswahl</param>
' <param name="viewMode">Anzeigeart (0: Detailansicht, 1: Vorschauansicht, 2: Eigenschaften, 3: Liste, 4: Miniaturansicht, 5: Große Symbole, 6: Kleine Symbole)</param>
' <returns>String (bei Mehfachauswahl sind die Dateien durch chr(9) getrennt)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function SelectFile(Optional ByVal InitialDir As String = vbNullString, _
                           Optional ByVal DlgTitle As String = SELECTBOX_FILE_DIALOG_TITLE, _
                           Optional ByVal FilterString As String = "Alle Dateien (*.*)", _
                           Optional ByVal MultiSelectEnabled As Boolean = False, _
                           Optional ByVal ViewMode As Long = -1) As String

    SelectFile = WizHook_GetFileName(InitialDir, DlgTitle, SELECTBOX_OPENTITLE, FilterString, MultiSelectEnabled, , ViewMode, False)

End Function

'---------------------------------------------------------------------------------------
' Function: SelectFolder
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Auswahldialog zur Verzeichnisauswahl
' </summary>
' <param name="InitDir">Startverzeichnis</param>
' <param name="DlgTitle">Dialogtitel</param>
' <param name="FilterString">Filterwerten - Beispiel: "(*.*)" oder "Alle (*.*)|Textdateien (*.txt)|Bilder (*.png;*.jpg;*.gif)</param>
' <param name="MultiSelect">Mehrfachauswahl</param>
' <param name="viewMode">Anzeigeart (0: Detailansicht, 1: Vorschauansicht, 2: Eigenschaften, 3: Liste, 4: Miniaturansicht, 5: Große Symbole, 6: Kleine Symbole)</param>
' <returns>String (bei Mehfachauswahl sind die Dateien durch chr(9) getrennt)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function SelectFolder(Optional ByVal InitialDir As String = vbNullString, _
                             Optional ByVal DlgTitle As String = SELECTBOX_FOLDER_DIALOG_TITLE, _
                             Optional ByVal FilterString As String = "*", _
                             Optional ByVal MultiSelectEnabled As Boolean = False, _
                             Optional ByVal ViewMode As Long = -1) As String

   SelectFolder = WizHook_GetFileName(InitialDir, DlgTitle, SELECTBOX_OPENTITLE, FilterString, MultiSelectEnabled, , ViewMode, True)

End Function

Private Function WizHook_GetFileName( _
                           ByVal InitialDir As String, _
                           ByVal DlgTitle As String, _
                           ByVal OpenTitle As String, _
                           ByVal FilterString As String, _
                           Optional ByVal MultiSelectEnabled As Boolean = False, _
                           Optional ByVal SplitDelimiter As String = "|", _
                           Optional ByVal ViewMode As Long = -1, _
                           Optional ByVal SelectFolderFlag As Boolean = False, _
                           Optional ByVal AppName As String) As String

'Zusammenfassung der Parameter von WizHook.GetFileName: http://www.team-moeller.de/?Tipps_und_Tricks:Wizhook-Objekt:GetFileName
'View  0: Detailansicht
'      1: Vorschauansicht
'      2: Eigenschaften
'      3: Liste
'      4: Miniaturansicht
'      5: Große Symbole
'      6: Kleine Symbole

'flags 4: Set Current Dir
'      8: Mehrfachauswahl möglich
'     32: Ordnerauswahldialog
'     64: Wert im Parameter "View" berücksichtigen

   Dim SelectedFileString As String
   Dim WizHookRetVal As Long

   If InStr(1, InitialDir, " ") > 0 Then
      InitialDir = """" & InitialDir & """"
   End If

   Dim Flags As Long
   Flags = 0
   If MultiSelectEnabled Then Flags = Flags + 8
   If SelectFolderFlag Then Flags = Flags + 32

   If ViewMode >= 0 Then
      Flags = Flags + 64
   Else
      ViewMode = 0
   End If

   WizHook.Key = 51488399
   WizHookRetVal = WizHook.GetFileName( _
                        Access.Application.hWndAccessApp, AppName, DlgTitle, OpenTitle, _
                        SelectedFileString, InitialDir, FilterString, 0, ViewMode, Flags, True)
   If WizHookRetVal = 0 Then
      If MultiSelectEnabled Then SelectedFileString = Replace(SelectedFileString, vbTab, SplitDelimiter)
      WizHook_GetFileName = SelectedFileString
   End If

End Function

'---------------------------------------------------------------------------------------
' Function: UNCPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt den UNC-Pfad zurück
' </summary>
' <param name="Path">Pfadangabe</param>
' <param name="IgnoreErrors">Fehler von API ignorieren</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function UncPath(ByVal Path As String, Optional ByVal IgnoreErrors As Boolean = True) As String
   
   Dim UNC As String * 512
   
   If VBA.Len(Path) = 1 Then Path = Path & ":"
   
   If WNetGetConnection(VBA.Left$(Path, 2), UNC, VBA.Len(UNC)) Then
   
      ' API-Routine gibt Fehler zurück:
      If IgnoreErrors Then
         UncPath = Path
      Else
         Err.Raise 5 ' Invalid procedure call or argument
      End If
   
   Else
   
      ' Ergebnis zurückgeben:
      UncPath = VBA.Left$(UNC, VBA.InStr(UNC, vbNullChar) - 1) & VBA.Mid$(Path, 3)
   
   End If
   
End Function

'---------------------------------------------------------------------------------------
' Property: TempPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Temp-Verzeichnis ermitteln
' </summary>
' <returns>String</returns>
' <remarks>
' Verwendet API GetTempPathA
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get TempPath() As String

   Dim TempString As String

   TempString = Space$(PATHLEN_MAX)
   API_GetTempPath PATHLEN_MAX, TempString
   TempString = Left$(TempString, InStr(TempString, Chr$(0)) - 1)
   If Len(TempString) = 0 Then
      TempString = DEFAULT_TEMPPATH_NOENV
   End If
   TempPath = TempString

End Property

Public Function GetNewTempFileName(Optional ByVal PathToUse As String = "", _
                         Optional ByVal FilePrefix As String = "", _
                         Optional ByVal FileExtension As String = "") As String

   Dim NewTempFileName As String
   
   If Len(PathToUse) = 0 Then
      PathToUse = TempPath
   End If

   NewTempFileName = String$(PATHLEN_MAX, 0)
   Call API_GetTempFilename(PathToUse, FilePrefix, 0&, NewTempFileName)

   NewTempFileName = Left$(NewTempFileName, InStr(NewTempFileName, Chr$(0)) - 1)

   'Datei wieder löschen, da nur Name benötigt wird
   Call Kill(NewTempFileName)

   If Len(FileExtension) > 0 Then 'Fileextension umschreiben
     NewTempFileName = Left$(NewTempFileName, Len(NewTempFileName) - 3) & FileExtension
   End If
   'eigentlich müsste man hier prüfen, ob Datei vorhanden ist.
   
   GetNewTempFileName = NewTempFileName

End Function

'---------------------------------------------------------------------------------------
' Function: ShortenFileName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dateipfad auf n Zeichen kürzen
' </summary>
' <param name="FullFileName">Vollständiger Pfad</param>
' <param name="MaxLen">gewünschte Länge</param>
' <returns>String</returns>
' <remarks>
' Hilfreich für die Anzeigen in schmalen Textfeldern \n
' Beispiel: <source>C:\Programme\...\Verzeichnis\Dateiname.txt</source>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ShortenFileName(ByVal FullFileName As Variant, ByVal MaxLen As Long) As String

   Dim FileString As String
   Dim Temp As String
   Dim TrimPos As Long

   FileString = Nz(FullFileName, vbNullString)
   If Len(FileString) <= MaxLen Then
      ShortenFileName = FileString
      Exit Function
   End If

   TrimPos = InStrRev(FileString, "\")
   Temp = Mid$(FileString, TrimPos)
   FileString = Left$(FileString, TrimPos - 1)

   TrimPos = MaxLen - Len(Temp) - 3
   If TrimPos < 2 Then
      FileString = "..." & Temp
   Else
      TrimPos = TrimPos \ 2
      FileString = Left$(FileString, TrimPos) & "..." & Right$(FileString, TrimPos) & Temp
   End If

   ShortenFileName = FileString

End Function

'---------------------------------------------------------------------------------------
' Function: FileNameWithoutPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dateinamen aus vollständiger Pfadangabe extrahieren
' </summary>
' <param name="FullPath">Dateiname inkl. Verzeichnis</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function FileNameWithoutPath(ByVal FullPath As Variant) As String

   Dim Temp As String
   Dim Pos As Long

   Temp = Nz(FullPath, vbNullString)
   Pos = InStrRev(Temp, "\")
   If Pos > 0 Then
      FileNameWithoutPath = Mid$(Temp, Pos + 1)
   Else
      FileNameWithoutPath = Temp
   End If

End Function

'---------------------------------------------------------------------------------------
' Function: CreateDirectory
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erstelle ein Verzeichnis inkl. aller fehlenden übergeordneten Verzeichnisse
' </summary>
' <param name="FullPath">Zu erstellendes Verzeichnis</param>
' <returns>Boolean: True = Verzeichnis wurde erstellt</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateDirectory(ByVal FullPath As String) As Boolean

   Dim PathBefore As String

   If Right$(FullPath, 1) = "\" Then
      FullPath = VBA.Left$(FullPath, Len(FullPath) - 1)
   End If

   If DirExists(FullPath) Then 'Verzeichnis ist bereits vorhanden
      CreateDirectory = False
      Exit Function
   End If

   PathBefore = VBA.Mid$(FullPath, 1, VBA.InStrRev(FullPath, "\") - 1)
   If Not DirExists(PathBefore) Then
      If CreateDirectory(PathBefore) = False Then
         CreateDirectory = False
         Exit Function
      End If
   End If

   VBA.MkDir FullPath

   CreateDirectory = True

End Function

Public Sub CreateDirectoryIfMissing(ByVal FullPath As String)
   CreateDirectory FullPath
End Sub

'---------------------------------------------------------------------------------------
' Function: FileExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft Existens einer Datei
' </summary>
' <param name="FullPath">Vollständige Pfadangabe</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function FileExists(ByVal FullPath As String) As Boolean

   Do While VBA.Right$(FullPath, 1) = "\"
      FullPath = VBA.Left$(FullPath, Len(FullPath) - 1)
   Loop

   FileExists = (VBA.Len(VBA.Dir$(FullPath, vbReadOnly Or vbHidden Or vbSystem)) > 0) And (VBA.Len(FullPath) > 0)
   VBA.Dir$ "\" ' Problemvermeidung: issue #109

End Function

'---------------------------------------------------------------------------------------
' Function: DirExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft Existenz eines Verzeichnisses
' </summary>
' <param name="FullPath">Vollständige Pfadangabe</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DirExists(ByVal FullPath As String) As Boolean

   If VBA.Right$(FullPath, 1) <> "\" Then
      FullPath = FullPath & "\"
   End If

   DirExists = (VBA.Dir$(FullPath, vbDirectory Or vbReadOnly Or vbHidden Or vbSystem) = ".")
   VBA.Dir$ "\" ' Problemvermeidung: issue #109
   
End Function

'---------------------------------------------------------------------------------------
' Function: GetFileUpdateDate
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Letztes Änderungsdatum einer Datei
' </summary>
' <param name="FullFileName">Vollständige Pfadangabe</param>
' <returns>Variant</returns>
' <remarks>
' Fehler von API-Funktion werden ignoriert
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFileUpdateDate(ByVal FullFileName As String) As Variant
   If FileExists(FullFileName) Then
      On Error Resume Next
      GetFileUpdateDate = FileDateTime(FullFileName)
   Else
      GetFileUpdateDate = Null
   End If
End Function

'---------------------------------------------------------------------------------------
' Function: ConvertStringToFileName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt aus einer Zeichenkette einen Dateinamen (ersetzt Sonderzeichen)
' </summary>
' <param name="Text">Ausgangsstring für Dateinamen</param>
' <param name="ReplaceWith">Zeichen als Ersatz für Sonderzeichen</param>
' <param name="CharsToReplace">Zeichen die mit ReplaceWith ersetzt werden</param>
' <param name="CharsToDelete">Zeichen die entfernt werden</param>
' <returns>String</returns>
' <remarks>
' Sonderzeichen: ? * " / ' : ( )
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ConvertStringToFileName(ByVal Text As String, _
                                   Optional ByVal ReplaceWith As String = "_", _
                                   Optional ByVal CharsToReplace As String = "/':()", _
                                   Optional ByVal CharsToDelete As String = "?*""") As String

   Dim FileName As String
   Dim i As Long

   FileName = Trim$(Text)

   For i = 1 To Len(CharsToDelete)
      FileName = Replace(FileName, Mid(CharsToReplace, i, 1), vbNullString)
   Next

   For i = 1 To Len(CharsToReplace)
      FileName = Replace(FileName, Mid(CharsToReplace, i, 1), ReplaceWith)
   Next

   ConvertStringToFileName = FileName

End Function

'---------------------------------------------------------------------------------------
' Function: GetFullPathFromRelativPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erezugt aus relativer Pfadangabe und "Basisverzeichnis" eine vollständige Pfadangabe
' </summary>
' <param name="RelativPath">relativer Pfad</param>
' <param name="BaseDir">Ausgangsverzeichnis</param>
' <returns>String</returns>
' <remarks>
' Beispiel:
' GetFullPathFromRelativPath("..\..\Test.txt", "C:\Programme\xxx\") => "C:\test.txt"
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFullPathFromRelativPath(ByVal RelativPath As String, _
                                           ByVal BaseDir As String) As String

   Dim FullPath As String
   Dim Pos As Long

   If Right$(BaseDir, 1) = "\" Then
      BaseDir = Left$(BaseDir, Len(BaseDir) - 1)
   End If

   FullPath = RelativPath
   If Mid$(FullPath, 2, 1) = ":" Or Left$(FullPath, 2) = "\\" Then ' absolut path !!!
      GetFullPathFromRelativPath = FullPath
      Exit Function
   ElseIf Left$(FullPath, 1) = "\" Then 'first dir
      Pos = InStr(3, BaseDir, "\")
      If Pos > 0 Then
         BaseDir = Left$(BaseDir, Pos - 1)
      End If
      GetFullPathFromRelativPath = BaseDir & FullPath
      Exit Function
   ElseIf FullPath = "." Then
      GetFullPathFromRelativPath = BaseDir
      Exit Function
   ElseIf Left$(FullPath, 2) = ".\" Then
      FullPath = Mid$(FullPath, 3)
   End If

   Do While Left$(FullPath, 3) = "..\"
      FullPath = Mid$(FullPath, 4)
      Pos = InStrRev(BaseDir, "\")
      If Pos > 0 Then
         BaseDir = Left$(BaseDir, Pos - 1)
      End If
   Loop

   GetFullPathFromRelativPath = BaseDir & "\" & FullPath

End Function

'---------------------------------------------------------------------------------------
' Function: GetRelativPathFromFullPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt einen relativen Pfad aus vollständiger Pfadangabe und Ausgangsverzeichnis
' </summary>
' <param name="FullPath">vollständiger Pfadangabe</param>
' <param name="BaseDir">Ausgangsverzeichnis</param>
' <param name="RelativePrefix">".\" als Kennung für relativen Pfad ergänzen</param>
' <returns>String</returns>
' <remarks>
' Beispiel:
' <code>
' GetRelativPathFromFullPath("C:\test.txt", "C:\Programme\xxx\", True)
' => ".\..\..\test.txt"
' </code>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetRelativPathFromFullPath(ByVal FullPath As String, _
                                           ByVal BaseDir As String, _
                                  Optional ByVal EnableRelativePrefix As Boolean = False, _
                                  Optional ByVal DisableDecreaseBaseDir As Boolean = False) As String

   Dim RelativPath As String
   
   If FullPath = BaseDir Then
      GetRelativPathFromFullPath = "."
      Exit Function
   End If

   If Right$(BaseDir, 1) <> "\" Then BaseDir = BaseDir & "\"
   If FullPath = BaseDir Then
      GetRelativPathFromFullPath = "."
      Exit Function
   End If
   
   If Not DisableDecreaseBaseDir Then
      RelativPath = TryGetRelativPathWithDecreaseBaseDir(FullPath, BaseDir, EnableRelativePrefix)
   Else
      RelativPath = FullPath
      If Right$(BaseDir, 1) <> "\" Then BaseDir = BaseDir & "\"
      If Len(BaseDir) > 0 Then
         If Nz(InStr(1, FullPath, BaseDir, vbTextCompare), 0) > 0 Then
            RelativPath = Mid$(FullPath, Len(BaseDir) + 1)
            If EnableRelativePrefix Then
               RelativPath = ".\" & RelativPath
            End If
         End If
      End If
   End If
   
   GetRelativPathFromFullPath = RelativPath

End Function

Private Function TryGetRelativPathWithDecreaseBaseDir(ByVal FullPath As String, ByVal BaseDir As String, ByVal EnableRelativePrefix As Boolean) As String

   Dim RelativPath As String
   Dim DecreaseCounter As Long
   Dim Pos As Long
   Dim i As Long
   
   RelativPath = BaseDir

   Do While InStr(1, FullPath, RelativPath) = 0
      Pos = InStrRev(Left$(RelativPath, Len(RelativPath) - 1), "\")
      RelativPath = Left$(RelativPath, Pos)
      DecreaseCounter = DecreaseCounter + 1
      If Len(RelativPath) = 0 Then
         DecreaseCounter = 0
         Exit Do
      End If
   Loop
   
   If Len(RelativPath) > 0 Then
      RelativPath = Replace(FullPath, RelativPath, vbNullString)
      For i = 1 To DecreaseCounter
         RelativPath = "..\" & RelativPath
      Next

      If EnableRelativePrefix Then
         RelativPath = ".\" & RelativPath
      End If
   Else
      RelativPath = FullPath
   End If

   TryGetRelativPathWithDecreaseBaseDir = RelativPath

End Function

'---------------------------------------------------------------------------------------
' Function: GetDirFromFullFileName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittels aus vollständer Pfadangabe einer Datei das Verzeichnis
' </summary>
' <param name="FullFileName">vollständer Pfadangabe</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetDirFromFullFileName(ByVal FullFileName As String) As String
   GetDirFromFullFileName = PathFromFullFileName(FullFileName)
End Function

Public Function PathFromFullFileName(ByVal FullFileName As Variant) As String

   Dim DirPath As String
   Dim Pos As Long

   DirPath = FullFileName
   Pos = InStrRev(DirPath, "\")
   If Pos > 0 Then
      DirPath = Left$(DirPath, Pos)
   Else
      DirPath = vbNullString
   End If

   PathFromFullFileName = DirPath

End Function

'---------------------------------------------------------------------------------------
' Sub: AddToZipFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei an Zip-Datei anhängen.
' </summary>
' <param name="ZipFile">Zip-Datei</param>
' <param name="FullFileName">Datei, die angehängt werden soll</param>
' <remarks>
' CreateObject("Shell.Application").Namespace(zipFile & "").CopyHere sFile & ""
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub AddToZipFile(ByVal ZipFile As String, ByVal FullFileName As String)

   If Not FileExists(ZipFile) Then
      CreateZipFile ZipFile
   End If

   With CreateObject("Shell.Application")
      .NameSpace(ZipFile & "").CopyHere FullFileName & ""
   End With

End Sub

'---------------------------------------------------------------------------------------
' Function: ExtractFromZipFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei aus Zip-Datei extrahieren
' </summary>
' <param name="ZipFile">Zip-Datei</param>
' <param name="Destination">Zielverzeichnis</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExtractFromZipFile(ByVal ZipFile As String, ByVal Destination As String) As String

   With CreateObject("Shell.Application")
      .NameSpace(Destination & "").CopyHere .NameSpace(ZipFile & "").Items
      ExtractFromZipFile = .NameSpace(ZipFile & "").Items.Item(0).Name
   End With

End Function

'---------------------------------------------------------------------------------------
' Function: CreateZipFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt leere Zipdatei
' </summary>
' <param name="ZipFile">Zip-Datei</param>
' <param name="DeleteExistingFile">Vorhandene Zip-Datei löschen</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateZipFile(ByVal ZipFile As String, Optional ByRef DeleteExistingFile As Boolean = False) As Boolean

   Dim FileHandle As Long

   If FileExists(ZipFile) Then
      If DeleteExistingFile Then
         Kill ZipFile
      Else
         CreateZipFile = False
         Exit Function
      End If
   End If

   FileHandle = FreeFile
   Open ZipFile For Output As #FileHandle
   Print #FileHandle, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String$(18, 0)
   Close #FileHandle

   CreateZipFile = FileExists(ZipFile)

End Function

'---------------------------------------------------------------------------------------
' Function: GetFileExtension
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt die Dateiendung einer Datei oder eines Pfads zurück.
' </summary>
' <param name="filePath">Dateipfad oder Dateiname</param>
' <returns>Dateiendung inkl. Trennzeichen</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFileExtension(ByVal FilePath As String, Optional ByVal WithDotBeforeExtension As Boolean = False) As String
   GetFileExtension = VBA.Strings.Mid$(FilePath, VBA.Strings.InStrRev(FilePath, ".") + (1 - Abs(WithDotBeforeExtension)))
End Function

Public Function OpenFile(FileName As String, Optional ByVal ReadOnlyMode As Boolean = False) As Boolean

   Dim strFile As String

   strFile = FileName
   If Len(Dir(strFile)) = 0 Then
      Err.Raise vbObjectError, "OpenFile", "Die Datei '" & FileName & vbNewLine & "' " & _
                  "konnte nicht gefunden werden." & vbNewLine & _
                  "Bitte überprüfen Sie den Datei-Pfad."
            Exit Function
   End If

   OpenFile = ShellExecute(strFile, "open")
   
End Function

Public Function OpenFilePath(FilePath As String) As Boolean

   Dim strFile As String

   strFile = FilePath
   If Len(Dir(FilePath, vbDirectory)) = 0 Then
      Err.Raise vbObjectError, "OpenFilePath", "Das Verzeichnis '" & FilePath & vbNewLine & "' " & _
                  "konnte nicht gefunden werden." & vbNewLine & _
                  "Bitte überprüfen Sie den Pfad."
            Exit Function
   End If

   OpenFilePath = ShellExecute(strFile, "open")
   
End Function

Private Function ShellExecute(ByVal FilePath As String, _
               Optional ByVal ApiOperation As String = vbNullString) As Boolean

   Dim Ret As Long
   Dim Directory As String
   Dim DeskWin As Long
   
   If Len(FilePath) = 0 Then
      ShellExecute = False
      Exit Function
   Else
      DeskWin = Application.hWndAccessApp
      Ret = API_ShellExecuteA(DeskWin, ApiOperation, FilePath, vbNullString, vbNullString, vbNormalFocus)
   End If
   
   If Ret = SE_ERR_NOTFOUND Then
      'Datei nicht gefunden
      MsgBox "Datei nicht gefunden" & vbNewLine & vbNewLine & _
             FilePath
      ShellExecute = False
      Exit Function
   ElseIf Ret = SE_ERR_NOASSOC Then
      ShellExecute = False
      Exit Function
' ToDo: "Öffnen mit"-Dialog verwenden:
      'Wenn die Dateierweiterung noch nicht bekannt ist...
      'wird der "Öffnen mit..."-Dialog angezeigt.
'      Directory = Space$(260)
'      Ret = GetSystemDirectory(Directory, Len(Directory))
'      Directory = Left$(Directory, Ret)
'      Call ShellExecuteA(DeskWin, vbNullString, "RUNDLL32.EXE", "shell32.dll, OpenAs_RunDLL " & _
'         FilePath, Directory, vbNormalFocus)
   End If
   
   ShellExecute = True

End Function
