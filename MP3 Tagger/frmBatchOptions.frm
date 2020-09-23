VERSION 5.00
Begin VB.Form frmBatchOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Options"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   Icon            =   "frmBatchOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkWritePictures 
      Caption         =   "&Write Pictures"
      Height          =   255
      Left            =   5040
      TabIndex        =   38
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CheckBox chkWriteLyrics 
      Caption         =   "&Write Lyrics"
      Height          =   255
      Left            =   5040
      TabIndex        =   37
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox chkID3V111ToID3V230 
      Caption         =   "ID3V1 V1.1 > ID3V2 V3.0"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CheckBox chkID3V230ToID3V111 
      Caption         =   "ID3V2 V3.0 > ID3V1 V1.1"
      Height          =   255
      Left            =   2640
      TabIndex        =   36
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "ID3 Tag Rewriting Options - Overrides"
      Height          =   3495
      Left            =   5040
      TabIndex        =   41
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtCommentOverride 
         Height          =   765
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   2520
         Width           =   3855
      End
      Begin VB.ComboBox cboGenreOverride 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtYearOverride 
         Height          =   285
         Left            =   600
         TabIndex        =   28
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtTrackOverride 
         Height          =   285
         Left            =   600
         TabIndex        =   26
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtSongOverride 
         Height          =   285
         Left            =   600
         TabIndex        =   24
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtAlbumOverride 
         Height          =   285
         Left            =   600
         TabIndex        =   22
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtArtistOverride 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Comment"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Genre"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Year"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "&Track"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "&Song"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "A&lbum"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "&Artist"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID3 Tag Rewriting Options"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtSeparator 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Text            =   " - "
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   600
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "U"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtAlbum 
         Height          =   285
         Left            =   600
         MaxLength       =   1
         TabIndex        =   6
         Text            =   "D"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtSong 
         Height          =   285
         Left            =   600
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "2"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtTrack 
         Height          =   285
         Left            =   600
         MaxLength       =   1
         TabIndex        =   12
         Text            =   "1"
         Top             =   2400
         Width           =   375
      End
      Begin VB.CheckBox chkRewriteID3V230 
         Caption         =   "Rewrite ID3V&2 V3.0 Tag"
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CheckBox chkRewriteID3V111 
         Caption         =   "Rewrite ID3V&1 V1.1 Tag"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Separator"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label lblInfo 
         Caption         =   "%1-%9 = Index Item"
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label lblTrack 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   2400
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label lblAlbum 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblSong 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblArtist 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Artist"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "A&lbum"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Song"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Track"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label lblExampleFileName 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8880
      TabIndex        =   40
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7800
      TabIndex        =   39
      Top             =   3720
      Width           =   975
   End
   Begin VB.CheckBox chkRemoveID3V230 
      Caption         =   "Remove ID3V&2 V3.0 Tag"
      Height          =   255
      Left            =   2640
      TabIndex        =   34
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CheckBox chkRemoveID3V111 
      Caption         =   "Remove ID3V&1 V1.1 Tag"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   3600
      Width           =   2175
   End
End
Attribute VB_Name = "frmBatchOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fstrDirs() As String
Private fstrItems() As String

Private Sub chkID3V111ToID3V230_Click()
    If chkID3V111ToID3V230.Value = 1 Then chkID3V230ToID3V111.Value = 0
End Sub

Private Sub chkID3V230ToID3V111_Click()
    If chkID3V230ToID3V111.Value = 1 Then chkID3V111ToID3V230.Value = 0
End Sub

Private Sub chkWriteLyrics_Click()
    If chkWriteLyrics.Value = 1 Then chkRewriteID3V230.Value = 1
End Sub

Private Sub chkWritePictures_Click()
    If chkWritePictures.Value = 1 Then chkRewriteID3V230.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    With typBatchOptions
        .bolCancelled = False
        .bolRemoveID3V111 = IIf(chkRemoveID3V111.Value = 1, True, False)
        .bolRemoveID3V230 = IIf(chkRemoveID3V230.Value = 1, True, False)
        .bolRewriteID3V111 = IIf(chkRewriteID3V111.Value = 1, True, False)
        .bolRewriteID3V230 = IIf(chkRewriteID3V230.Value = 1, True, False)
        .strArtist = txtArtist.Text
        .strAlbum = txtAlbum.Text
        .strSong = txtSong.Text
        .strTrack = txtTrack.Text
        .strSeparator = txtSeparator.Text
        .strArtistOverride = txtArtistOverride.Text
        .strAlbumOverride = txtAlbumOverride.Text
        .strSongOverride = txtSongOverride.Text
        If Len(txtTrackOverride.Text) <> 0 Then
            .intTrackOverride = IIf(IsNumeric(txtTrackOverride.Text) = True, CInt(txtTrackOverride.Text), 0)
            If .intTrackOverride < 0 Then .intTrackOverride = -1
        Else
            .intTrackOverride = -1
        End If
        .strCommentOverride = txtCommentOverride.Text
        .strYearOverride = txtYearOverride.Text
        .strGenreOverride = cboGenreOverride.Text
        .bolCopyID3V111ToID3V230 = IIf(chkID3V111ToID3V230.Value = 1, True, False)
        .bolCopyID3V230ToID3V111 = IIf(chkID3V230ToID3V111.Value = 1, True, False)
        .bolWriteLyrics = IIf(chkWriteLyrics.Value = 1, True, False)
        .bolWritePictures = IIf(chkWritePictures.Value = 1, True, False)
    End With
    
    Me.Hide
End Sub

Private Sub Form_Load()
    lblExampleFileName.Caption = typExample.strDir & typExample.strFile
    typExample.strFile = Left(typExample.strFile, Len(typExample.strFile) - 4)
    lblInfo.Caption = "1-9 = Index Item" & vbCrLf & "D - Current Dir" & vbCrLf & "U - Up One Dir"
    typBatchOptions.bolCancelled = True
    fstrDirs = Split(typExample.strDir, "\")
    fstrItems = Split(typExample.strFile, txtSeparator.Text)
    LoadGenres
    txtArtist_Change
    txtAlbum_Change
    txtTrack_Change
    txtSong_Change
End Sub

Private Sub txtAlbum_Change()
    Dim intNumber As Integer
    
    On Error Resume Next
    
    lblAlbum.Caption = ""
    If Len(txtAlbum.Text) <> 0 Then
        Select Case LCase(txtAlbum.Text)
            Case "d"
                lblAlbum.Caption = fstrDirs(UBound(fstrDirs) - 1)
            Case "u"
                lblAlbum.Caption = fstrDirs(UBound(fstrDirs) - 2)
            Case Else
                If IsNumeric(txtAlbum.Text) = True Then
                    intNumber = CInt(txtAlbum.Text)
                    lblAlbum.Caption = fstrItems(intNumber - 1)
                End If
        End Select
    End If
End Sub

Private Sub txtAlbum_GotFocus()
    txtAlbum.SelStart = 0
    txtAlbum.SelLength = 1
End Sub

Private Sub txtAlbum_KeyPress(KeyAscii As Integer)
    txtAlbum.SelStart = 0
    txtAlbum.SelLength = 1
End Sub

Private Sub txtAlbumOverride_GotFocus()
    txtAlbumOverride.SelStart = 0
    txtAlbumOverride.SelLength = Len(txtAlbumOverride.Text)
End Sub

Private Sub txtArtist_Change()
    Dim intNumber As Integer
    
    On Error Resume Next
    
    lblArtist.Caption = ""
    If Len(txtArtist.Text) <> 0 Then
        Select Case LCase(txtArtist.Text)
            Case "d"
                lblArtist.Caption = fstrDirs(UBound(fstrDirs) - 1)
            Case "u"
                lblArtist.Caption = fstrDirs(UBound(fstrDirs) - 2)
            Case Else
                If IsNumeric(txtArtist.Text) = True Then
                    intNumber = CInt(txtArtist.Text)
                    lblArtist.Caption = fstrItems(intNumber - 1)
                End If
        End Select
    End If
End Sub

Private Sub txtArtist_GotFocus()
    txtArtist.SelStart = 0
    txtArtist.SelLength = 1
End Sub

Private Sub txtArtist_KeyPress(KeyAscii As Integer)
    txtArtist.SelStart = 0
    txtArtist.SelLength = 1
End Sub

Private Sub txtArtistOverride_GotFocus()
    txtArtistOverride.SelStart = 0
    txtArtistOverride.SelLength = Len(txtArtistOverride.Text)
End Sub

Private Sub txtCommentOverride_GotFocus()
    txtCommentOverride.SelStart = 0
    txtCommentOverride.SelLength = Len(txtCommentOverride.Text)
End Sub

Private Sub txtSeparator_Change()
    If Len(txtSeparator.Text) <> 0 Then fstrItems = Split(typExample.strFile, txtSeparator.Text)
End Sub

Private Sub txtSeparator_GotFocus()
    txtSeparator.SelStart = 0
    txtSeparator.SelLength = Len(txtSeparator.Text)
End Sub

Private Sub txtSong_Change()
    Dim intNumber As Integer
    
    On Error Resume Next
    
    lblSong.Caption = ""
    If Len(txtSong.Text) <> 0 Then
        Select Case LCase(txtSong.Text)
            Case "d"
                lblSong.Caption = fstrDirs(UBound(fstrDirs) - 1)
            Case "u"
                lblSong.Caption = fstrDirs(UBound(fstrDirs) - 2)
            Case Else
                If IsNumeric(txtSong.Text) = True Then
                    intNumber = CInt(txtSong.Text)
                    lblSong.Caption = fstrItems(intNumber - 1)
                End If
        End Select
    End If
End Sub

Private Sub txtSong_GotFocus()
    txtSong.SelStart = 0
    txtSong.SelLength = 1
End Sub

Private Sub txtSong_KeyPress(KeyAscii As Integer)
    txtSong.SelStart = 0
    txtSong.SelLength = 1
End Sub

Private Sub txtSongOverride_GotFocus()
    txtSongOverride.SelStart = 0
    txtSongOverride.SelLength = Len(txtSongOverride.Text)
End Sub

Private Sub txtTrack_Change()
    Dim intNumber As Integer
    
    On Error Resume Next
    
    lblTrack.Caption = ""
    If Len(txtTrack.Text) <> 0 Then
        Select Case LCase(txtTrack.Text)
            Case "d"
                lblTrack.Caption = fstrDirs(UBound(fstrDirs) - 1)
            Case "u"
                lblTrack.Caption = fstrDirs(UBound(fstrDirs) - 2)
            Case Else
                If IsNumeric(txtTrack.Text) = True Then
                    intNumber = CInt(txtTrack.Text)
                    lblTrack.Caption = fstrItems(intNumber - 1)
                End If
        End Select
    End If
End Sub

Private Sub txtTrack_GotFocus()
    txtTrack.SelStart = 0
    txtTrack.SelLength = 1
End Sub

Private Sub txtTrack_KeyPress(KeyAscii As Integer)
    txtTrack.SelStart = 0
    txtTrack.SelLength = 1
End Sub

Private Sub LoadGenres()
    Dim strGenres()
    Dim lngCounter As Long
    
    strGenres = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", "Hip-Hop", _
                       "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&B", "Rap", "Reggae", "Rock", _
                       "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", "Soundtrack", _
                       "Euro-Techno", "Ambient", "Trip-Hop", "Vocal", "Jazz+Funk", "Fusion", "Trance", _
                       "Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", "Noise", _
                       "AlternRock", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
                       "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno-Industrial", "Electronic", _
                       "Pop-Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", _
                       "Top 40", "Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", _
                       "New Wave", "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo-Fi", "Tribal", _
                       "Acid Punk", "Acid Jazz", "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", _
                       "Folk", "Folk/Rock", "National folk", "Swing", "Fast-fusion", "Bebob", "Latin", _
                       "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", "Progressive Rock", _
                       "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", _
                       "Easy Listening", "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", _
                       "Sonata", "Symphony", "Booty Bass", "Primus", "Porn Groove", "Satire", "Slow Jam", _
                       "Club", "Tango", "Samba", "Folklore", "Ballad", "Powder Ballad", "Rhythmic Soul", _
                       "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Capella", "Euro-House", "Dance Hall", _
                       "Goa", "Drum & Bass", "Club House", "Hardcore", "Terror", "Indie", "BritPop", _
                       "NegerPunk", "Polsk Punk", "Beat", "Christian Gangsta", "Heavy Metal", "Black Metal", _
                       "Crossover", "Contemporary C", "Christian Rock", "Merengue", "Salsa", "Thrash Metal", _
                       "Anime", "JPop", "SynthPop")
    cboGenreOverride.Clear
    cboGenreOverride.AddItem ""
    
    For lngCounter = 0 To UBound(strGenres)
        cboGenreOverride.AddItem strGenres(lngCounter)
    Next
End Sub

Private Sub txtTrackOverride_GotFocus()
    txtTrackOverride.SelStart = 0
    txtTrackOverride.SelLength = Len(txtTrackOverride.Text)
End Sub

Private Sub txtYearOverride_GotFocus()
    txtYearOverride.SelStart = 0
    txtYearOverride.SelLength = Len(txtYearOverride.Text)
End Sub
