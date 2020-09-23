Attribute VB_Name = "modMisc"
Option Explicit

Public Type BatchOptions
    bolCancelled As Boolean
    bolRemoveID3V111 As Boolean
    bolRemoveID3V230 As Boolean
    bolRewriteID3V111 As Boolean
    bolRewriteID3V230 As Boolean
    strArtist As String
    strAlbum As String
    strSong As String
    strTrack As String
    strSeparator As String
    strArtistOverride As String
    strAlbumOverride As String
    strSongOverride As String
    intTrackOverride As Integer
    strCommentOverride As String
    strYearOverride As String
    strGenreOverride As String
    bolCopyID3V111ToID3V230 As Boolean
    bolCopyID3V230ToID3V111 As Boolean
    bolWriteLyrics As Boolean
    bolWritePictures As Boolean
End Type

Public Type Example
    strDir As String
    strFile As String
End Type

Public typBatchOptions As BatchOptions
Public typExample As Example

Public Type MS
    M As Integer
    S As Integer
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WinAMP_GetTrackTime = 105
Private Const WM_USER = &H400

Public glngWinAmp As Long

Public Function AppendPath(ByVal strPath As String) As String
    On Error Resume Next
    
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    AppendPath = strPath
End Function

Public Function FileExists(ByVal strFileName As String) As Boolean
    Dim intFile As Integer
    
    On Error GoTo ErrHan
    
    intFile = FreeFile
    Open strFileName For Input As intFile
    Close intFile
    
    FileExists = True
    Exit Function
    
ErrHan:
    FileExists = False
End Function

Public Function TMS(ByVal lngSeconds As Long) As MS
    TMS.M = lngSeconds \ 60
    lngSeconds = lngSeconds - (TMS.M * 60)
    TMS.S = lngSeconds
End Function

Public Sub GetWinAMPWindow()
    glngWinAmp = FindWindow("Winamp v1.x", vbNullString)
End Sub

Public Function GetWinAMPPosition() As String
    Dim lngTime As Long
    Dim msTime As MS
    
    On Local Error GoTo ErrHan
    
    If glngWinAmp = 0 Then
        GetWinAMPWindow
        If glngWinAmp = 0 Then Exit Function
    End If
    
    lngTime = SendMessage(glngWinAmp, WM_USER, 0, WinAMP_GetTrackTime)
    lngTime = lngTime / 1000
    msTime = TMS(lngTime)
    GetWinAMPPosition = Format(msTime.M, "00") & ":" & Format(msTime.S, "00")
    
    Exit Function
    
ErrHan:
End Function
