Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10215
    DatasheetFontHeight =11
    ItemSuffix =53
    Left =9135
    Top =2910
    Right =19350
    Bottom =10830
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x212b6fd80e9ce340
    End
    Caption ="Access Code Library - Import Wizard"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
        End
        Begin Section
            Height =7937
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9690
                    Top =90
                    Width =397
                    Height =397
                    TabIndex =2
                    Name ="cmdSelectLocalRepository"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadad00000000000adada ,
                        0x003333333330adad0b03333333330ada0fb03333333330ad0bfb03333333330a ,
                        0x0fbfb000000000000bfbfbfbfb0adada0fbfbfbfbf0dadad0bfb0000000adada ,
                        0xa000adadadad000ddadadadadadad00aadadadad0dad0d0ddadadadad000dada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Lokales Root-Verzeichnis ausw�hlen"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =9688
                    Top =1050
                    Width =397
                    Height =397
                    TabIndex =5
                    Name ="cmdSelectFile"
                    Caption ="&Select File"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadad00000000000adada ,
                        0x003333333330adad0b03333333330ada0fb03333333330ad0bfb03333333330a ,
                        0x0fbfb000000000000bfbfbfbfb0adada0fbfbfbfbf0dadad0bfb0000000adada ,
                        0xa000adadadad000ddadadadadadad00aadadadad0dad0d0ddadadadad000dada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Datei(en) mit Auswahldialog anf�gen"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4702
                    Top =120
                    Width =4896
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    LeftMargin =29
                    Name ="txtLocalRepositoryPath"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =113
                            Top =120
                            Width =4575
                            Height =315
                            Name ="Label5"
                            Caption ="Lokales Root-Verzeichnis der Code-Bibliothek:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    AccessKey =65
                    IMESentenceMode =3
                    Left =113
                    Top =1075
                    Width =9021
                    Height =315
                    TabIndex =3
                    Name ="txtFileString"
                    StatusBarText ="Pfad ab Root-Verzeichnis"
                    OnKeyDown ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =735
                            Width =4410
                            Height =315
                            Name ="Label11"
                            Caption ="Datei &anf�gen"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9195
                    Top =1050
                    Width =397
                    Height =397
                    TabIndex =4
                    Name ="cmdAddFile"
                    Caption ="+"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Datei aus Textzeile anf�gen"
                End
                Begin ListBox
                    OverlapFlags =247
                    AccessKey =68
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =113
                    Top =1814
                    Width =9979
                    Height =3232
                    TabIndex =6
                    BoundColumn =1
                    Name ="lstImportFiles"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2835;3402"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    ControlTipText ="Liste der zu importierenden Dateien\015\012Markierte Eintr�ge k�nnen mit {Entf}-"
                        "Taste aus der Liste entfernt werden."
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =2280
                            Height =315
                            Name ="Label13"
                            Caption ="Ausgew�hlte &Dateien:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =73
                    Left =7035
                    Top =7155
                    Width =2205
                    Height =555
                    TabIndex =9
                    Name ="cmdImportFiles"
                    Caption ="Dateien &importieren"
                    OnClick ="[Event Procedure]"
                    OnMouseDown ="[Event Procedure]"
                End
                Begin OptionGroup
                    OverlapFlags =223
                    AccessKey =77
                    Left =113
                    Top =5220
                    Width =9982
                    Height =2616
                    TabIndex =7
                    Name ="ogImportMode"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =222
                            Top =5277
                            Width =6585
                            Height =315
                            Name ="lblImportMode"
                            Caption ="Import-&Modus"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =352
                            Top =6050
                            OptionValue =0
                            Name ="Option17"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =577
                                    Top =6020
                                    Width =6240
                                    Height =315
                                    Name ="Label18"
                                    Caption ="vorhandene Elemente nicht �berschreiben"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =352
                            Top =6440
                            OptionValue =2
                            Name ="Option19"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =582
                                    Top =6410
                                    Width =6225
                                    Height =315
                                    Name ="Label20"
                                    Caption ="vorhandene Elemente �berschreiben"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =352
                            Top =7295
                            OptionValue =1
                            Name ="Option21"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =577
                                    Top =7265
                                    Width =6240
                                    Height =315
                                    Name ="Label22"
                                    Caption ="nur die ausgew�hlte Datei(en) importieren"
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =285
                    Top =5670
                    Width =6525
                    Height =285
                    Name ="Label23"
                    Caption ="fehlende Abh�ngigkeiten automatisch erg�nzen"
                End
                Begin Label
                    OverlapFlags =215
                    Left =283
                    Top =6916
                    Width =6525
                    Height =285
                    Name ="Label24"
                    Caption ="keine Abh�ngigkeitspr�fung"
                End
                Begin CommandButton
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =7035
                    Top =5896
                    Width =2925
                    Height =555
                    TabIndex =8
                    Name ="Command25"
                    Caption ="Abh�ngigkeiten anzeigen"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7035
                    Top =6720
                    Width =2895
                    Height =315
                    FontWeight =700
                    ForeColor =8210719
                    Name ="labInfo"
                    Caption ="Dateien wurden importiert"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =0
                    Height =0
                    Name ="sysFirst"
                End
                Begin CheckBox
                    OverlapFlags =215
                    AccessKey =84
                    Left =7035
                    Top =5340
                    TabIndex =10
                    Name ="chkImportTests"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"
                    ControlTipText ="Die Testklassen der importierten Code-Modulen ebenfalls importieren"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =7311
                            Top =5280
                            Width =2730
                            Height =315
                            Name ="Bezeichnungsfeld32"
                            Caption ="inkl. &Test-Klassen"
                        End
                    End
                End
                Begin CommandButton
                    Transparent = NotDefault
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =10185
                    Width =29
                    Height =29
                    TabIndex =11
                    Name ="cmdClose"
                    Caption ="Schlie�en"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6979
                    Top =2136
                    Width =3120
                    Height =2910
                    TabIndex =12
                    BorderColor =0
                    Name ="txtCodeModuleDescription"
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6979
                    Top =1814
                    Width =3120
                    Height =285
                    TabIndex =13
                    BorderColor =0
                    Name ="txtCodeModuleName"
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =6975
                    Top =1500
                    Width =3105
                    Height =273
                    FontSize =8
                    TabIndex =14
                    ForeColor =4138256
                    Name ="tbViewCodeModuleDescription"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"
                    Caption ="Beschreibung anzeigen"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9357
                    Top =7143
                    Width =555
                    Height =555
                    TabIndex =15
                    Name ="cmdOpenMenu"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Weitere Aktionen ..."
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: ACLibImportWizardForm
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Wizard-Formular f�r Import der CodeLib-Elemente
' </summary>
' <remarks></remarks>
'\ingroup ACLibAddInImportWizard
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/ImportWizard/ACLibImportWizardForm.frm</file>
'  <description>Maske f�r Import-Wizard</description>
'  <use>base/modErrorHandler.bas</use>
'  <use>_codelib/addins/ImportWizard/defGlobal_ACLibImportWizard.bas</use>
'  <use>api/winapi/modWinAPI.bas</use>
'  <use>api/winapi/WinApiShortcutMenu.cls</use>
'  <use>file/modFiles.bas</use>
'  <use>data/dao/TempDbHandler.cls</use>
'  <use>data/modSQL_Tools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

' verwendete Erweiterungen
Private Const EXTENSION_KEY_AppFile As String = "AppFile"
Private Const APPFILE_PROPNAME_AppIcon As String = "AppIcon"

Private Const TEMPDB_TABNAME As String = "tRepositoryFiles"
Private Const TEMPDB_TABDDL As String = "create table " & TEMPDB_TABNAME & " (LocalRepositoryPath varchar(255) primary key, ObjectName varchar(255), Description memo)"
Private m_TempDb As TempDbHandler

Private m_LastSelectionID As Variant

Private Sub bindTextbox(ByRef tb As Textbox, Optional ByVal BaseFolderPath As String = vbNullString)

   'Latebindung, damit ApplicationHandler_DirTextbox-Klasse nicht vorhanden sein muss
   Dim objDirTextbox As Object ' = ApplicationHandler_DirTextbox

   'Standard-Instanz verwenden:
On Error GoTo HandleErr

   Set objDirTextbox = CurrentApplication.Extensions("DirTextbox")
   
   'Textbox binden
   If Not (objDirTextbox Is Nothing) Then
      Set objDirTextbox.Textbox = tb
      objDirTextbox.BaseFolderPath = BaseFolderPath
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "bindTextbox", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub chkImportTests_AfterUpdate()
   CurrentACLibConfiguration.ImportTestsDefaultValue = Nz(Me.chkImportTests.Value, False)
End Sub

Private Sub cmdAddFile_Click()

On Error GoTo HandleErr

   addFileFromFileName Me.txtFileString & ""

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdAddFile_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub cmdClose_Click()
   DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdImportFiles_Click()

   Dim fileNameArray() As String
   Dim ArraySize As Long, i As Long
   Dim lb As ListBox
   
On Error GoTo HandleErr

   Set lb = Me.lstImportFiles
   
   ArraySize = lb.ListCount
   
   If ArraySize <= 0 Then
      MsgBox "Es sind keine Dateien ausgew�hlt.", vbInformation
      Exit Sub
   End If
   
   ReDim fileNameArray(ArraySize)
   
   For i = 0 To ArraySize - 1
      fileNameArray(i) = lb.ItemData(i)
   Next
   
   Me.labInfo.Caption = "Importvorgang l�uft ..."
   Me.labInfo.Visible = True
   Me.Repaint
   
   CurrentACLibFileManager.ImportRepositoryFiles fileNameArray, Nz(Me.ogImportMode.Value, 0), Nz(Me.chkImportTests.Value, False)
   
   Me.labInfo.Caption = "Dateien wurden importiert"
   Me.Repaint
   
   TempDb.Execute "delete from " & TEMPDB_TABNAME
   lb.Requery
   
   Me.SetFocus

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdImportFiles_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub cmdImportFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo HandleErr

   If Button = 2 Then
      OpenImportFileShortcutMenu
      Button = 0
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdImportFiles_MouseDown", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Function OpenImportFileShortcutMenu() As Long

   Dim mnu As WinApiShortcutMenu
On Error GoTo HandleErr

   Set mnu = New WinApiShortcutMenu

   With mnu
      Set .MenuControl = Me.cmdImportFiles
      Set .AccessForm = Me
      .ControlSection = acDetail
      
      .AddMenuItem 11, "Alle Elemente aus der Codebibliothek importieren", , 1
      .AddMenuItem 10, "Importieren", MF_POPUP, 1
      
      .AddMenuItem -1, "", MF_SEPARATOR
      
      .AddMenuItem 31, "Alle vorhandenen Elemente exportieren", , 2
      .AddMenuItem 32, "Alle vorhandenen Module exportieren", , 2
      .AddMenuItem 30, "Exportieren", MF_POPUP, 2
      
      .AddMenuItem -2, "", MF_SEPARATOR

      .AddMenuItem 21, "Alle vorhandenen Elemente aktualisieren", , 3
      .AddMenuItem 22, "Alle vorhandenen Module aktualisieren", , 3
      .AddMenuItem 20, "Aktualisieren", MF_POPUP, 3

      
   End With
   
   Select Case mnu.OpenMenu
      Case 11
         CurrentACLibFileManager.ImportAllFilesFromRepository Nz(Me.ogImportMode.Value, clim_ImportAllUsedItems)
      Case 21
         CurrentACLibFileManager.RefreshAll Nz(Me.ogImportMode.Value, clim_ImportMissingItems), True
      Case 22
         CurrentACLibFileManager.RefreshAllModules Nz(Me.ogImportMode.Value, clim_ImportMissingItems), True
      Case 31
         CurrentACLibFileManager.ExportAll
      Case 32
         CurrentACLibFileManager.ExportAllModules
      Case Else
         '
   End Select
   
   Set mnu = Nothing

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ShowImportFileShortcutMenu", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Sub cmdOpenMenu_Click()
  
On Error GoTo HandleErr

   OpenImportFileShortcutMenu

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdOpenMenu_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Sub

Private Sub cmdSelectFile_Click()
   
   Dim strStartFolder As String
   Dim strFiles As String
   Dim fileArray() As String
   Dim Pos As Long

On Error GoTo HandleErr

   strStartFolder = Replace(Me.txtFileString.Value & "", "/", "\")
   If Len(strStartFolder) > 0 Then
      Do While Left$(strStartFolder, 1) = "\"
         strStartFolder = Mid$(strStartFolder, 1)
         If Len(strStartFolder) = 0 Then Exit Do
      Loop
   End If
   
   strStartFolder = CurrentLocalRepositoryPath & strStartFolder
   Do While Not DirExists(strStartFolder)
      Pos = InStrRev(strStartFolder, "\")
      If Pos = 0 Then Exit Do
      strStartFolder = Left$(strStartFolder, Pos - 1)
   Loop

   Me.sysFirst.SetFocus
   Me.cmdSelectFile.SetFocus

   strFiles = SelectFile(strStartFolder, , , True)
   If Len(strFiles) = 0 Then
      Exit Sub
   End If
   
   fileArray = Split(strFiles, "|")
   addFiles fileArray

   

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdSelectFile_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub addFiles(ByRef fileArray() As String)
   
   Dim lb As ListBox
   Dim i As Long
   Dim ArraySize As Long
   
   Dim cli As CodeLibInfo
   
On Error GoTo HandleErr

   ArraySize = UBound(fileArray)
   
   Set lb = Me.lstImportFiles
   For i = 0 To ArraySize
      cli = CurrentACLibFileManager.GetCodeLibInfoFromFilePath(fileArray(i))
      TempDb.Execute "insert into " & TEMPDB_TABNAME & " (ObjectName, LocalRepositoryPath, Description) VALUES (" & _
                           SqlTools.TextToSqlText(cli.Name) & ", " & SqlTools.TextToSqlText(getLocalRepositoryPath(fileArray(i))) & _
                           ", " & SqlTools.TextToSqlText(cli.Description) & ")", dbFailOnError
   Next
   
   lb.Requery
   Me.labInfo.Visible = (lb.ListCount = 0)
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdSelectFile_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Function getLocalRepositoryPath(ByRef FullPath As String) As String

On Error GoTo HandleErr

   getLocalRepositoryPath = Replace(GetRelativPathFromFullPath(Replace(FullPath, "/", "\"), CurrentLocalRepositoryPath, False), "\", "/")

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "getLocalRepositoryPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Property Get CurrentLocalRepositoryPath() As String

On Error GoTo HandleErr

   CurrentLocalRepositoryPath = Me.txtLocalRepositoryPath.Value

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentLocalRepositoryPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Sub cmdSelectLocalRepository_Click()
   
   Dim selectedRepositoryPath As String
   
On Error GoTo HandleErr

   selectedRepositoryPath = SelectFolder(Nz(Me.txtLocalRepositoryPath, vbNullString), "Lokalen Repository-Ordner ausw�hlen", , False, 1)
   
   If Len(selectedRepositoryPath) > 0 Then
      If Right$(selectedRepositoryPath, 1) = "\" Then
         selectedRepositoryPath = Left$(selectedRepositoryPath, Len(selectedRepositoryPath) - 1)
      End If
      
      setLocalRepositoryPath selectedRepositoryPath
      
   End If
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "cmdSelectLocalRepository_Click", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub setEnableMode()

On Error GoTo HandleErr

   Me.cmdImportFiles.Enabled = Len(Me.txtLocalRepositoryPath.Value & vbNullString) > 0

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "setEnableMode", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub Form_Load()

On Error GoTo HandleErr

   Me.Caption = CurrentApplication.ApplicationTitle & "  " & ChrW(&H25AA) & "  Version " & CurrentApplication.Version
   loadIconFromAppFiles
   
   Me.txtLocalRepositoryPath.Value = CurrentACLibConfiguration.LocalRepositoryPath
   Me.chkImportTests.Value = CurrentACLibConfiguration.ImportTestsDefaultValue
   
   EnableCodeModuleDescription Me.tbViewCodeModuleDescription.Value
   
   setEnableMode
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "Form_Load", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub Form_Resize()
   Me.cmdImportFiles.Top = Me.InsideHeight - Me.lblImportMode.Left - Me.cmdImportFiles.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   If Not (m_TempDb Is Nothing) Then
      Me.lstImportFiles.RowSource = vbNullString
      DBEngine.Idle dbRefreshCache
      m_TempDb.Dispose
   End If
   DisposeCurrentApplicationHandler
End Sub

Private Sub lstImportFiles_AfterUpdate()
   RefreshCodeModuleDescription
End Sub

Private Sub lstImportFiles_DblClick(Cancel As Integer)
   OpenSelectItemFormImportFilesListboxInTextViewer
End Sub

Private Sub lstImportFiles_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo HandleErr

   If KeyCode = vbKeyDelete Then
      removeSelectedItemsFromListbox
   ElseIf KeyCode = vbKeyF2 Then
      OpenSelectItemFormImportFilesListboxInTextViewer
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "lstImportFiles_KeyDown", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub removeSelectedItemsFromListbox()

   Dim lb As ListBox
   Dim selectedItem As Variant
   Dim strItemFilter As String

On Error GoTo HandleErr

   Set lb = Me.lstImportFiles
   
   For Each selectedItem In lb.ItemsSelected
      strItemFilter = ", " & SqlTools.TextToSqlText(lb.Column(1, selectedItem))
   Next
   
   If Len(strItemFilter) <= 2 Then
      Exit Sub
   End If
   
   strItemFilter = Mid$(strItemFilter, 3)
   TempDb.Execute "delete from " & TEMPDB_TABNAME & " where LocalRepositoryPath IN (" & strItemFilter & ")"
   
   lb.Requery
   
   RefreshCodeModuleDescription

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "removeSelectedItemsFormListbox", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub lstImportFiles_LostFocus()
On Error Resume Next
   m_LastSelectionID = Me.lstImportFiles.Column(1)
   Me.lstImportFiles = Null
End Sub

Private Sub tbViewCodeModuleDescription_AfterUpdate()
   EnableCodeModuleDescription Me.tbViewCodeModuleDescription.Value
   If Len(m_LastSelectionID) > 0 Then
      SelectListItem m_LastSelectionID
   End If
End Sub

Private Sub SelectListItem(ByVal sItemID As String)
   Dim i As Long
   Dim lb As ListBox
   Set lb = Me.lstImportFiles
   For i = 0 To (lb.ListCount - 1)
      If lb.Column(1, i) = sItemID Then
         lb.SetFocus
         lb.Selected(i) = True
         RefreshCodeModuleDescriptionFromID sItemID, lb.Column(0, i)
         Exit Sub
      End If
   Next
End Sub

Private Sub txtFileString_GotFocus()
On Error Resume Next
   bindTextbox Me.txtFileString, CurrentLocalRepositoryPath
End Sub

Private Sub txtFileString_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo HandleErr

   If KeyCode = vbKeyReturn Then
      If Me.txtFileString.Text = ".." Then
         Exit Sub
      ElseIf Replace(Right$(Me.txtFileString.Text, 3), "/", "\") = "\.." Then
         Exit Sub
      Else
         addFileFromFileName Me.txtFileString.Text
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyF2 Then
      OpenRepositoryFileInTextViewer Me.txtFileString.Text
      KeyCode = 0
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "txtFileString_KeyDown", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub addFileFromFileName(ByVal FileString As String)

On Error GoTo HandleErr

   FileString = Trim$(Replace(FileString, "/", "\"))
   If Len(FileString) > 0 Then
      Do While Left$(FileString, 1) = "\"
         FileString = Trim$(Mid$(FileString, 2))
         If Len(FileString) = 0 Then Exit Sub
      Loop
      addFile CurrentLocalRepositoryPath & FileString
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "addFileFromFileName", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub


Private Function addFile(ByRef newFileName As String) As Boolean

   Dim fileArray(0) As String
   
On Error GoTo HandleErr

   If Len(newFileName) = 0 Then
      Exit Function
   End If
   
   newFileName = Replace(newFileName, "\\", "\")

   If Not FileExists(newFileName) Then
      MsgBox "Diese Datei ist nicht vorhanden", vbInformation
      addFile = False
      Exit Function
   End If
   
   fileArray(0) = newFileName
   addFiles fileArray
   
   addFile = True

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "addFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Sub txtLocalRepositoryPath_AfterUpdate()

On Error GoTo HandleErr

   setLocalRepositoryPath Me.txtLocalRepositoryPath & vbNullString

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "txtLocalRepositoryPath_AfterUpdate", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub setLocalRepositoryPath(ByRef newRoot As String)

On Error GoTo HandleErr

   CurrentACLibConfiguration.LocalRepositoryPath = newRoot
   
   'damit m�gliche Modifikationen aus CurrentACLibConfiguration �bernommen werden:
   Me.txtLocalRepositoryPath.Value = CurrentACLibConfiguration.LocalRepositoryPath
   
   setEnableMode

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "setLocalRepositoryPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub txtLocalRepositoryPath_BeforeUpdate(Cancel As Integer)

   Dim strNewPath As String
On Error GoTo HandleErr

   strNewPath = Me.txtLocalRepositoryPath & ""
   
   If Len(strNewPath) > 0 Then
      Cancel = Not DirExists(strNewPath)
      If Cancel Then
         MsgBox "Verzeichnis ist nicht vorhanden", vbInformation
      End If
   End If
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "txtLocalRepositoryPath_BeforeUpdate", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub loadIconFromAppFiles()

   Dim strIconFilePath As String
   Dim strIconFileName As String
   
   'Latebindung, damit ApplicationHandler_AppFile-Klasse nicht vorhanden sein muss
   Dim objAppFile As Object ' ... ApplicationHandler_AppFile

On Error GoTo HandleErr

   If Val(SysCmd(acSysCmdAccessVer)) <= 9 Then 'Abbruch, da Ac00 sonst abst�rzt
      Exit Sub
   End If

   Set objAppFile = CurrentApplication.Extensions(EXTENSION_KEY_AppFile)
   
   'Textbox binden
   If Not (objAppFile Is Nothing) Then
      strIconFileName = ACLibIconFileName
      strIconFilePath = CurrentACLibConfiguration.ACLibConfigDirectory

      If Len(ACLibIconFileName) = 0 Then 'nur Temp-Datei erzeugen
         strIconFileName = Me.Name & ".ico"
         strIconFilePath = TempPath
      End If
      
      strIconFilePath = strIconFilePath & strIconFileName
      
      If Len(Dir$(strIconFilePath)) = 0 Then
         If Not objAppFile.CreateAppFile(APPFILE_PROPNAME_AppIcon, strIconFilePath) Then
            Exit Sub
         End If
      End If
      
      WinAPI.Image.SetFormIconFromFile Me, strIconFilePath
      
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "loadIconFromAppFiles", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Property Get TempDb() As TempDbHandler

On Error GoTo HandleErr

   If m_TempDb Is Nothing Then
      Set m_TempDb = New TempDbHandler
      m_TempDb.CreateDatabase True, True
      m_TempDb.CreateTable TEMPDB_TABNAME, TEMPDB_TABDDL
      Me.lstImportFiles.RowSource = "select ObjectName, LocalRepositoryPath FROM [" & m_TempDb.CurrentDatabase.Name & "]." & TEMPDB_TABNAME
   End If
   Set TempDb = m_TempDb

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "TempDb", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Sub OpenSelectItemFormImportFilesListboxInTextViewer()
   OpenRepositoryFileInTextViewer Me.lstImportFiles.Column(1)
End Sub

Private Sub OpenRepositoryFileInTextViewer(ByVal sRelativeFilePath As String)
   Dim FullPath As String
   FullPath = CurrentACLibFileManager.GetRepositoryFullPath(sRelativeFilePath)
   WinAPI.Shell.Execute FullPath, "open"
End Sub

Private Sub EnableCodeModuleDescription(ByVal bViewDescription As Boolean)

   With Me.lstImportFiles
      If bViewDescription Then
         .Width = Me.lblImportMode.Left + Me.lblImportMode.Width - .Left
         RefreshCodeModuleDescription
      Else
         .Width = Me.ogImportMode.Width
      End If
   End With
   
   Me.txtCodeModuleDescription.Visible = bViewDescription
   Me.txtCodeModuleName.Visible = bViewDescription
   
End Sub

Private Sub RefreshCodeModuleDescription()
   RefreshCodeModuleDescriptionFromID Nz(Me.lstImportFiles.Column(1), vbNullString), Nz(Me.lstImportFiles.Column(0), vbNullString)
End Sub

Private Sub RefreshCodeModuleDescriptionFromID(ByVal sLocalRepositoryPath As String, ByVal sName As String)

   Dim strDescription As String
   If Len(sLocalRepositoryPath) > 0 Then
      strDescription = Nz(m_TempDb.LookupSQL("select Description from " & TEMPDB_TABNAME & " where LocalRepositoryPath = " & SqlTools.TextToSqlText(sLocalRepositoryPath)), vbNullString)
   End If
   Me.txtCodeModuleName.Value = sName
   Me.txtCodeModuleDescription.Value = strDescription
   
End Sub
