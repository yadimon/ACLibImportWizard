Attribute VB_Name = "FileTools"
Attribute VB_Description = "Funktionen f�r Dateioperationen"
'---------------------------------------------------------------------------------------
' Module: FileTools
'---------------------------------------------------------------------------------------
'/**
'\author    Josef Poetzl
'\short     Funktionen f�r Dateioperationen
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

Private Const m_SELECTBOX_File_DlgTitle As String = "Datei ausw�hlen"
Private Const m_SELECTBOX_Folder_DlgTitle As String = "Ordner ausw�hlen"
Private Const m_SELECTBOX_OpenTitle As String = "ausw�hlen"

Private Const m_DEFAULT_TEMPPATH_NoEnv As String = "C:\"
Private Const m_MAXPATHLEN As Long = 255

#If VBA7 Then

Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private Declare PtrSafe Function API_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

#Else

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private Declare Function API_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

#End If

'---------------------------------------------------------------------------------------
' Function: SelectFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei mittels Dialog ausw�hlen
' </summary>
' <param name="InitDir">Startverzeichnis</param>
' <param name="DlgTitle">Dialogtitel</param>
' <param name="FilterString">Filterwerten - Beispiel: "(*.*)" oder "Alle (*.*)|Textdateien (*.txt)|Bilder (*.png;*.jpg;*.gif)</param>
' <param name="MultiSelect">Mehrfachauswahl</param>
' <param name="viewMode">Anzeigeart (0: Detailansicht, 1: Vorschauansicht, 2: Eigenschaften, 3: Liste, 4: Miniaturansicht, 5: Gro�e Symbole, 6: Kleine Symbole)</param>
' <returns>String (bei Mehfachauswahl sind die Dateien durch chr(9) getrennt)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function SelectFile(Optional ByVal InitialDir As String = vbNullString, _
                           Optional ByVal DlgTitle As String = m_SELECTBOX_File_DlgTitle, _
                           Optional ByVal FilterString As String = "Alle Dateien (*.*)", _
                           Optional ByVal MultiSelectEnabled As Boolean = False, _
                           Optional ByVal ViewMode As Long = -1) As String

    SelectFile = WizHook_GetFileName(InitialDir, DlgTitle, m_SELECTBOX_OpenTitle, FilterString, MultiSelectEnabled, , ViewMode, False)

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
' <param name="viewMode">Anzeigeart (0: Detailansicht, 1: Vorschauansicht, 2: Eigenschaften, 3: Liste, 4: Miniaturansicht, 5: Gro�e Symbole, 6: Kleine Symbole)</param>
' <returns>String (bei Mehfachauswahl sind die Dateien durch chr(9) getrennt)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function SelectFolder(Optional ByVal InitialDir As String = vbNullString, _
                             Optional ByVal DlgTitle As String = m_SELECTBOX_Folder_DlgTitle, _
                             Optional ByVal FilterString As String = "*", _
                             Optional ByVal MultiSelectEnabled As Boolean = False, _
                             Optional ByVal ViewMode As Long = -1) As String

   SelectFolder = WizHook_GetFileName(InitialDir, DlgTitle, m_SELECTBOX_OpenTitle, FilterString, MultiSelectEnabled, , ViewMode, True)

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
'      5: Gro�e Symbole
'      6: Kleine Symbole

'flags 4: Set Current Dir
'      8: Mehrfachauswahl m�glich
'     32: Ordnerauswahldialog
'     64: Wert im Parameter "View" ber�cksichtigen

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
' Gibt den UNC-Pfad zur�ck
' </summary>
' <param name="Path">Pfadangabe</param>
' <param name="IgnoreErrors">Fehler von API ignorieren</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function UNCPath(ByVal Path As String, Optional ByVal IgnoreErrors As Boolean = True) As String

  Dim UNC As String * 512

  If Len(Path) = 1 Then Path = Path & ":"

  If WNetGetConnection(Left$(Path, 2), UNC, Len(UNC)) Then

    ' API-Routine gibt Fehler zur�ck:
    If IgnoreErrors Then
      UNCPath = Path
    Else
      Err.Raise 5 ' Invalid procedure call or argument
    End If

  Else

    ' Ergebnis zur�ckgeben:
    UNCPath = Left$(UNC, InStr(UNC, vbNullChar) - 1) _
            & Mid$(Path, 3)

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

   TempString = Space$(m_MAXPATHLEN)
   API_GetTempPath m_MAXPATHLEN, TempString
   TempString = Left$(TempString, InStr(TempString, Chr$(0)) - 1)
   If Len(TempString) = 0 Then
      TempString = m_DEFAULT_TEMPPATH_NoEnv
   End If
   TempPath = TempString

End Property

'---------------------------------------------------------------------------------------
' Function: ShortenFileName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dateipfad auf n Zeichen k�rzen
' </summary>
' <param name="FullFileName">Vollst�ndiger Pfad</param>
' <param name="MaxLen">gew�nschte L�nge</param>
' <returns>String</returns>
' <remarks>
' Hilfreich f�r die Anzeigen in schmalen Textfeldern \n
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
' Dateinamen aus vollst�ndiger Pfadangabe extrahieren
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
' Erstelle ein Verzeichnis inkl. aller fehlenden �bergeordneten Verzeichnisse
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
      FullPath = Left$(FullPath, Len(FullPath) - 1)
   End If

   If Len(Dir$(FullPath, vbDirectory)) > 0 Then 'Verzeichnis ist bereits vorhanden
      CreateDirectory = False
      Exit Function
   End If

   PathBefore = Mid$(FullPath, 1, InStrRev(FullPath, "\") - 1)
   If Len(Dir$(PathBefore, vbDirectory)) = 0 Then
      If CreateDirectory(PathBefore) = False Then
         CreateDirectory = False
         Exit Function
      End If
   End If

   MkDir FullPath

   CreateDirectory = True

End Function

'---------------------------------------------------------------------------------------
' Function: FileExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Pr�ft Existens einer Datei
' </summary>
' <param name="FullPath">Vollst�ndige Pfadangabe</param>
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
      '6 = vbNormal or vbHidden or vbSystem

End Function

'---------------------------------------------------------------------------------------
' Function: DirExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Pr�ft Existenz eines Verzeichnisses
' </summary>
' <param name="FullPath">Vollst�ndige Pfadangabe</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DirExists(ByVal FullPath As String) As Boolean

   If Right$(FullPath, 1) <> "\" Then
      FullPath = FullPath & "\"
   End If

   DirExists = (Dir$(FullPath, vbDirectory Or vbReadOnly Or vbHidden Or vbSystem) = ".")

End Function

'---------------------------------------------------------------------------------------
' Function: GetFileUpdateDate
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Letztes �nderungsdatum einer Datei
' </summary>
' <param name="FullFileName">Vollst�ndige Pfadangabe</param>
' <returns>Variant</returns>
' <remarks>
' Fehler von API-Funktion werden ignoriert
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFileUpdateDate(ByVal FullFileName As String) As Variant
   If Len(Dir$(FullFileName)) > 0 Then
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
' <param name="Text">Ausgangsstring f�r Dateinamen</param>
' <param name="ReplaceWith">Zeichen als Ersatz f�r Sonderzeichen</param>
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

   Dim fileName As String
   Dim i As Long
   
   fileName = Trim$(Text)
   
   For i = 1 To Len(CharsToDelete)
      fileName = Replace(fileName, Mid(CharsToReplace, i, 1), vbNullString)
   Next
   
   For i = 1 To Len(CharsToReplace)
      fileName = Replace(fileName, Mid(CharsToReplace, i, 1), ReplaceWith)
   Next
   
   ConvertStringToFileName = fileName

End Function

'---------------------------------------------------------------------------------------
' Function: GetFullPathFromRelativPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erezugt aus relativer Pfadangabe und "Basisverzeichnis" eine vollst�ndige Pfadangabe
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
' Erzeugt einen relativen Pfad aus vollst�ndiger Pfadangabe und Ausgangsverzeichnis
' </summary>
' <param name="FullPath">vollst�ndiger Pfadangabe</param>
' <param name="BaseDir">Ausgangsverzeichnis</param>
' <param name="RelativePrefix">".\" als Kennung f�r relativen Pfad erg�nzen</param>
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
                                  Optional ByVal EnableRelativePrefix As Boolean = False) As String

   Dim RelativPath As String
   Dim Pos As Long
   Dim Counter As Long, i As Long

   If FullPath = BaseDir Then
      GetRelativPathFromFullPath = "."
      Exit Function
   End If

   If Right$(BaseDir, 1) <> "\" Then BaseDir = BaseDir & "\"
   If FullPath = BaseDir Then
      GetRelativPathFromFullPath = "."
      Exit Function
   End If

   RelativPath = BaseDir

   Do While InStr(1, FullPath, RelativPath) = 0
      Pos = InStrRev(Left$(RelativPath, Len(RelativPath) - 1), "\")
      RelativPath = Left$(RelativPath, Pos)
      Counter = Counter + 1
      If Len(RelativPath) = 0 Then
         Counter = 0
         Exit Do
      End If
   Loop

   If Len(RelativPath) > 0 Then
      RelativPath = Replace(FullPath, RelativPath, vbNullString)
      For i = 1 To Counter
         RelativPath = "..\" & RelativPath
      Next

      If EnableRelativePrefix Then
         RelativPath = ".\" & RelativPath
      End If

   Else
      RelativPath = FullPath
   End If

   GetRelativPathFromFullPath = RelativPath

End Function

'---------------------------------------------------------------------------------------
' Function: GetDirFromFullFileName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittels aus vollst�nder Pfadangabe einer Datei das Verzeichnis
' </summary>
' <param name="FullFileName">vollst�nder Pfadangabe</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetDirFromFullFileName(ByVal FullFileName As String) As String

   Dim DirPath As String
   Dim Pos As Long

   DirPath = FullFileName
   Pos = InStrRev(DirPath, "\")
   If Pos > 0 Then
      DirPath = Left$(DirPath, Pos)
   Else
      DirPath = vbNullString
   End If

   GetDirFromFullFileName = DirPath

End Function

'---------------------------------------------------------------------------------------
' Sub: AddToZipFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei an Zip-Datei anh�ngen.
' </summary>
' <param name="ZipFile">Zip-Datei</param>
' <param name="FullFileName">Datei, die angeh�ngt werden soll</param>
' <remarks>
' CreateObject("Shell.Application").Namespace(zipFile & "").CopyHere sFile & ""
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub AddToZipFile(ByVal ZipFile As String, ByVal FullFileName As String)

   If Len(Dir$(ZipFile)) = 0 Then
      CreateZipFile ZipFile
   End If

   With CreateObject("Shell.Application")
      .Namespace(ZipFile & "").CopyHere FullFileName & ""
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
      .Namespace(Destination & "").CopyHere .Namespace(ZipFile & "").Items
      ExtractFromZipFile = .Namespace(ZipFile & "").Items.item(0).Name
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
' <param name="DeleteExistingFile">Vorhandene Zip-Datei l�schen</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateZipFile(ByVal ZipFile As String, Optional DeleteExistingFile As Boolean = False) As Boolean

   Dim fileHandle As Long

   If Len(Dir$(ZipFile)) > 0 Then
      If DeleteExistingFile Then
         Kill ZipFile
      Else
         CreateZipFile = False
         Exit Function
      End If
   End If
   
   fileHandle = FreeFile
   Open ZipFile For Output As #fileHandle
   Print #fileHandle, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String$(18, 0)
   Close #fileHandle

   CreateZipFile = (Len(Dir$(ZipFile)) > 0)

End Function

'---------------------------------------------------------------------------------------
' Function: GetFileExtension
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt die Dateiendung einer Datei oder eines Pfads zur�ck.
' </summary>
' <param name="filePath">Dateipfad oder Dateiname</param>
' <returns>Dateiendung inkl. Trennzeichen</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFileExtension(ByVal FilePath As String) As String
    GetFileExtension = VBA.Strings.Mid$(FilePath, VBA.Strings.InStrRev(FilePath, "."))
End Function
