VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_AppFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Erweiterung f�r ApplicationHandler-Klasse: Anwendungsspezifische Dateien verwalten"
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_File (Josef P�tzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung f�r ApplicationHandler-Klasse: Anwendungsspezifische Dateien verwalten
' </summary>
' <remarks>
' Aktivierung im _config_Application-Modul:\n
' <code>modApplication.AddApplicationHandlerExtension New ApplicationHandler_AppFile</code>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler_AppFile.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>api/winapi/WinApiFileInfo.cls</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const m_EXTENSION_KEY As String = "AppFile"

Private Const m_EXTENSION_PROPERTY_AppFileTableName As String = "AppFileTableName"
Private Const m_TABLE_AppFiles As String = "usys_AppFiles"
Private Const m_TABLE_FIELD_id As String = "id"
Private Const m_TABLE_FIELD_version As String = "version"
Private Const m_TABLE_FIELD_file As String = "file"


'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef objRef As ApplicationHandler)
   Set m_ApplicationHandler = objRef
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = m_EXTENSION_KEY
End Property

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, ByRef exists As Boolean)
   If sExtensionKey = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ByRef ExtensionReference As Object)
   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal sExtensionKey As String, ByVal PropertyName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   If sExtensionKey = m_EXTENSION_KEY Then
      getExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If

End Sub

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
   Set m_ApplicationHandler = Nothing
End Sub



'---------------------------------------------------------------------------------------
' Erg�nzungen f�r Ereiterung: ApplicationHandler_AppFile
'---------------------------------------------------------------------------------------
'

Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, _
                                       ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                       ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed

   Select Case PropertyName
      Case m_EXTENSION_PROPERTY_AppFileTableName
         ResumeMessage = m_TABLE_AppFiles
         
      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error
         
   End Select

End Sub

Public Property Get AppFileTableName() As String
   AppFileTableName = m_TABLE_AppFiles
End Property

'---------------------------------------------------------------------------------------
' Function: CreateAppFile (2009-07-30)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei aus usys_AppFiles extrahieren
' </summary>
' <param name="Param"></param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateAppFile(ByVal sFileID As String, ByVal sFileName As String) As Boolean

   Dim f As Integer
   Dim BinFile() As Byte
   Dim lngFieldSize As Long
   Dim rst As DAO.Recordset
   Dim fld As DAO.Field

   Set rst = CodeDb.OpenRecordset("select file from " & m_TABLE_AppFiles & " where id='" & sFileID & "'")
   If rst.EOF Then
      CreateAppFile = False
   Else
      Set fld = rst.Fields(0)
      lngFieldSize = fld.FieldSize
      If lngFieldSize > 0 Then

         ReDim BinFile(lngFieldSize - 1)
         BinFile = fld.GetChunk(0, lngFieldSize)

        ' Datei erstellen
         f = FreeFile
         Open sFileName For Binary As #f
         Put #f, , BinFile()
         Close #f

         CreateAppFile = True
      Else
         CreateAppFile = False
      End If

   End If

End Function

'---------------------------------------------------------------------------------------
' Function: SaveAppFile
'---------------------------------------------------------------------------------------
'/
' <summary>
' Datei in usys_AppFiles speichern
' </summary>
' <param name="sFileID">Kennung in Tabelle (Feld "id")</param>
' <param name="sFileName">Dateiname</param>
' <param name="saveVersion">Version abspeichern</param>
' <remarks>
' </remarks>
'/
'---------------------------------------------------------------------------------------
Public Function SaveAppFile(ByVal FileId As String, ByVal fileName As String, _
                   Optional ByVal SaveVersion As Boolean = False) As Boolean

   Dim f As Integer
   Dim BinFile() As Byte
   Dim ArraySize As Long
   Dim strVersion As String
   Dim FileInfo As WinApiFileInfo
   
On Error GoTo HandleErr

   f = FreeFile
   Open fileName For Binary As #f
   ArraySize = LOF(f) - 1
   If ArraySize >= 0 Then
      ReDim BinFile(ArraySize)
      Get #f, , BinFile()
   End If
   Close #f

   If ArraySize < 0 Then
      SaveAppFile = False
      Exit Function
   End If

   Dim rst As DAO.Recordset
   Set rst = CodeDb.OpenRecordset("select id, file, Version from " & m_TABLE_AppFiles & " where id='" & FileId & "'")
   If rst.EOF Then
      rst.AddNew
      rst.Fields("id") = FileId
   Else
      rst.Edit
   End If
   rst.Fields("file").AppendChunk BinFile
   If SaveVersion Then
      Set FileInfo = New WinApiFileInfo
      strVersion = FileInfo.GetFileVersion(fileName)
          Set FileInfo = Nothing
      If Len(strVersion) = 0 Then
         strVersion = Format$(Nz(FileDateTime(fileName), vbNullString), "yyyy.mm.dd")
      End If
      rst.Fields("version") = strVersion
   End If
   rst.Update
   rst.Close
   Set rst = Nothing
   
   SaveAppFile = True

ExitHere:
   Exit Function
   
HandleErr:
   Dim tabCreateErrCnt As Long
   Dim lngErrNumber As Long, strErrDescription As String

   lngErrNumber = Err.Number
   strErrDescription = Err.Description

   If lngErrNumber = 3078 And tabCreateErrCnt = 0 Then
      tabCreateErrCnt = 1
      lngErrNumber = lngErrNumber
      Err.Clear
      If createAppFileTable Then
         Resume
      End If
   End If

   Err.Raise lngErrNumber, "ExtAppFile.SaveAppFile", strErrDescription

End Function

Private Function createAppFileTable() As Boolean
   
   Dim strSQL As String

   strSQL = "create table " & m_TABLE_AppFiles & " ( " & m_TABLE_FIELD_id & " varchar(20) primary key, " & _
             m_TABLE_FIELD_version & " varchar(10), " & m_TABLE_FIELD_file & " image)"
   CodeDb.Execute strSQL, dbFailOnError

   createAppFileTable = (Err.Number = 0)

End Function


'---------------------------------------------------------------------------------------
' Ereignisbehandlung von m_ApplicationHandler
'
' AppFileBeforeCreateFile
Private Sub m_ApplicationHandler_AppFileBeforeCreateFile(ByVal sFileID As String, ByVal sFileName As String, _
                              ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   Dim bolCompleted As Boolean

   bolCompleted = CreateAppFile(sFileID, sFileName)
   If bolCompleted Then
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed
   Else 'Fehler r�ckmelden
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
   End If
   
End Sub

' AppFileBeforeSaveFile
Private Sub m_ApplicationHandler_AppFileBeforeSaveFile(ByVal sFileID As String, ByVal sFileName As String, _
                              ByVal SaveVersion As Boolean, ByRef ResumeMode As ApplicationHandlerResumeModes, _
                              ByRef ResumeMessage As Variant)

   Dim bolCompleted As Boolean

   bolCompleted = SaveAppFile(sFileID, sFileName, SaveVersion)
   If bolCompleted Then
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed
   Else 'Fehler r�ckmelden
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
   End If
   
End Sub
