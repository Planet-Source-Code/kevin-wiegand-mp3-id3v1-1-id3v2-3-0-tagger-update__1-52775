VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "MP3 Tagger Version 2.0"
   ClientHeight    =   9000
   ClientLeft      =   510
   ClientTop       =   600
   ClientWidth     =   12000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrWinAmp 
      Interval        =   250
      Left            =   1200
      Top             =   720
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   600
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraID3 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   66
      Top             =   8160
      Width           =   11775
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
         Height          =   375
         Left            =   10800
         TabIndex        =   72
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdCopyToID3V230 
         Caption         =   "&Copy to ID3V2 3.0"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   61
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdBatchOptions 
         Caption         =   "&Batch Options"
         Height          =   375
         Left            =   6120
         TabIndex        =   62
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdCopyToID3V11 
         Caption         =   "&Copy to ID3V1 1.1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   60
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveID3 
         Caption         =   "&Remove ID3"
         Height          =   375
         Left            =   1200
         TabIndex        =   59
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveID3 
         Caption         =   "&Save ID3"
         Height          =   375
         Left            =   0
         TabIndex        =   58
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame fraID3Buttons 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   64
      Top             =   4440
      Width           =   11775
      Begin VB.CheckBox chkID3V230 
         Caption         =   "ID3V&2 3.0"
         Height          =   315
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CheckBox chkID3V11 
         Caption         =   "ID3V&1 1.1"
         Height          =   315
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8400
      TabIndex        =   63
      Top             =   720
      Width           =   3495
      Begin VB.CommandButton cmdAddDirectory 
         Caption         =   "+"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         ToolTipText     =   "Add To List of Directories"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdRemoveDirectory 
         Caption         =   "-"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         ToolTipText     =   "Remove From List of Directories"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdScanDirectory 
         Caption         =   "&Scan"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         ToolTipText     =   "Read The Contents of the Directory"
         Top             =   0
         Width           =   735
      End
      Begin VB.CheckBox chkRecursiveSearch 
         Caption         =   "&Recursive Search"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   0
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   67
      Top             =   8625
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20638
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMP3 
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5530
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cboDirectories 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   11775
   End
   Begin VB.Frame fraInfo 
      Caption         =   "ID3 Information"
      Height          =   3135
      Left            =   120
      TabIndex        =   65
      Top             =   4920
      Width           =   11775
      Begin VB.CheckBox chkInfoSyncLyrics 
         Caption         =   "&Sync Lyrics"
         Height          =   315
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkInfoPictures 
         Caption         =   "&Pictures"
         Height          =   315
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkInfoLyrics 
         Caption         =   "&Lyrics"
         Height          =   315
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkInfoID3V2 
         Caption         =   "&ID3V2"
         Height          =   315
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkInfoGeneral 
         Caption         =   "&General"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame fraSyncLyrics 
         Height          =   2535
         Left            =   120
         TabIndex        =   73
         Top             =   480
         Visible         =   0   'False
         Width           =   11535
         Begin VB.CommandButton cmdCopyFromLyrics 
            Caption         =   "Lyrics>Sync"
            Height          =   375
            Left            =   10200
            TabIndex        =   78
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdInsertTime 
            Caption         =   "&Insert Time"
            Height          =   375
            Left            =   10200
            TabIndex        =   77
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveSyncLyrics 
            Caption         =   "&Save"
            Height          =   375
            Left            =   10200
            TabIndex        =   76
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtSyncLyrics 
            Height          =   2175
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   75
            ToolTipText     =   "Dbl-Click To Offset Times"
            Top             =   240
            Width           =   9975
         End
         Begin VB.CommandButton cmdLoadSyncLyrics 
            Caption         =   "&Load"
            Height          =   375
            Left            =   10200
            TabIndex        =   74
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10200
            TabIndex        =   79
            Top             =   240
            Width           =   1170
         End
      End
      Begin VB.Frame fraPictures 
         Height          =   2535
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   11535
         Begin VB.CommandButton cmdSetTitle 
            Caption         =   "&Set"
            Height          =   375
            Left            =   7080
            TabIndex        =   57
            Top             =   2040
            Width           =   735
         End
         Begin VB.ComboBox cboPictureType 
            Height          =   315
            ItemData        =   "frmMain.frx":030A
            Left            =   3480
            List            =   "frmMain.frx":030C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtPictureTitle 
            Height          =   285
            Left            =   3480
            TabIndex        =   56
            Top             =   2040
            Width           =   3495
         End
         Begin VB.CommandButton cmdRemovePicture 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   2400
            TabIndex        =   52
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton cmdSavePicture 
            Caption         =   "&Save"
            Height          =   375
            Left            =   2400
            TabIndex        =   51
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton cmdLoadPicture 
            Caption         =   "&Load"
            Height          =   375
            Left            =   2400
            TabIndex        =   50
            Top             =   1080
            Width           =   975
         End
         Begin VB.Image picPicture 
            Height          =   2175
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "&Picture Type"
            Height          =   195
            Left            =   3480
            TabIndex        =   53
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "&Picture Title"
            Height          =   195
            Left            =   3480
            TabIndex        =   55
            Top             =   1800
            Width           =   840
         End
      End
      Begin VB.Frame fraID3V2 
         Height          =   2535
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Visible         =   0   'False
         Width           =   11535
         Begin VB.TextBox txtUnknown 
            Height          =   525
            Left            =   7440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   1800
            Width           =   3855
         End
         Begin VB.CommandButton cmdDeleteURL 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   10320
            TabIndex        =   41
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton cmdEditURL 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   9240
            TabIndex        =   40
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton cmdAddURL 
            Caption         =   "&Add"
            Height          =   375
            Left            =   8160
            TabIndex        =   39
            Top             =   960
            Width           =   975
         End
         Begin VB.ListBox lstURL 
            Height          =   450
            Left            =   7440
            TabIndex        =   38
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox txtLanguage 
            Height          =   285
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   36
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox txtEncodedBy 
            Height          =   285
            Left            =   1080
            TabIndex        =   34
            Top             =   1680
            Width           =   5895
         End
         Begin VB.TextBox txtCopyright 
            Height          =   285
            Left            =   1080
            TabIndex        =   32
            Top             =   1320
            Width           =   5895
         End
         Begin VB.TextBox txtComposer 
            Height          =   285
            Left            =   1080
            TabIndex        =   30
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox txtOriginalArtist 
            Height          =   285
            Left            =   1080
            TabIndex        =   26
            Top             =   240
            Width           =   5895
         End
         Begin VB.TextBox txtSubtitle 
            Height          =   285
            Left            =   1080
            TabIndex        =   28
            Top             =   600
            Width           =   5895
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "&Unknown"
            Height          =   195
            Left            =   7440
            TabIndex        =   42
            Top             =   1560
            Width           =   690
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "&URL"
            Height          =   195
            Left            =   7440
            TabIndex        =   37
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "&Language"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "&Encoded By"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   870
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "&Copyright"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1320
            Width           =   660
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "&Composer"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "&Subtitle"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   525
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&Original Artist"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame fraGeneral 
         Height          =   2535
         Left            =   120
         TabIndex        =   68
         Top             =   480
         Visible         =   0   'False
         Width           =   11535
         Begin VB.ComboBox cboGenre 
            Height          =   315
            Left            =   600
            TabIndex        =   21
            Text            =   "cboGenre"
            Top             =   2040
            Width           =   3735
         End
         Begin VB.TextBox txtComment 
            Height          =   2085
            Left            =   7440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtYear 
            Height          =   285
            Left            =   600
            TabIndex        =   19
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtTrack 
            Height          =   285
            Left            =   600
            TabIndex        =   17
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox txtSong 
            Height          =   285
            Left            =   600
            TabIndex        =   15
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox txtAlbum 
            Height          =   285
            Left            =   600
            TabIndex        =   13
            Top             =   600
            Width           =   5895
         End
         Begin VB.TextBox txtArtist 
            Height          =   285
            Left            =   600
            TabIndex        =   11
            Top             =   240
            Width           =   5895
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "&Comment"
            Height          =   195
            Left            =   6720
            TabIndex        =   22
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "&Genre"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   435
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Year"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   330
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "&Track"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Song"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "A&lbum"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Artist"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   345
         End
      End
      Begin VB.Frame fraLyrics 
         Height          =   2535
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Visible         =   0   'False
         Width           =   11535
         Begin VB.CommandButton cmdCopyFromSync 
            Caption         =   "Sync>Lyrics"
            Height          =   375
            Left            =   10200
            TabIndex        =   80
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveLyrics 
            Caption         =   "&Save"
            Height          =   375
            Left            =   10200
            TabIndex        =   48
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtLyrics 
            Height          =   2175
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   47
            Top             =   240
            Width           =   9975
         End
         Begin VB.CommandButton cmdLoadLyrics 
            Caption         =   "&Load"
            Height          =   375
            Left            =   10200
            TabIndex        =   49
            Top             =   2040
            Width           =   1215
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Directory To Search:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Type Graphics
    strFileName As String
    bytType As Byte
    strTitle As String
    dblStartPosition As Double
    dblLength As Double
    bolExtracted As Boolean
End Type

Private Type Lyric
    strLine As String
    intMinute As Integer
    intSecond As Integer
End Type

Private Lyrics() As Lyric
Private fbolStop As Boolean
Private fbolStopBatch As Boolean
Private fGraphics() As Graphics

Private Sub barStatus_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If fbolStopBatch = False Then
        barStatus.Panels(1).Text = "Cancelling Request..."
        fbolStopBatch = True
    End If
End Sub

Private Sub cboPictureType_Click()
    Dim lngCounter As Long
    
    On Local Error Resume Next
    
    picPicture.Picture = LoadPicture()
    txtPictureTitle.Text = ""
    If cboPictureType.ListIndex = -1 Then Exit Sub
    
    For lngCounter = 1 To UBound(fGraphics)
        If fGraphics(lngCounter).bytType = cboPictureType.ItemData(cboPictureType.ListIndex) Then
            picPicture.Picture = LoadPicture(fGraphics(lngCounter).strFileName)
            txtPictureTitle.Text = fGraphics(lngCounter).strTitle
        End If
    Next
End Sub

Private Sub chkID3V11_Click()
    If chkID3V11.Value = 1 Then
        chkID3V230.Value = 0
    Else
        If chkID3V230.Value = 0 Then chkID3V11.Value = 1
    End If
    
    If chkID3V11.Value = 1 Then
        cmdCopyToID3V11.Enabled = False
        cmdCopyToID3V230.Enabled = True
        chkInfoID3V2.Enabled = False
        chkInfoLyrics.Enabled = False
        chkInfoSyncLyrics.Enabled = False
        chkInfoPictures.Enabled = False
        chkInfoGeneral.Value = 1
    Else
        cmdCopyToID3V11.Enabled = True
        cmdCopyToID3V230.Enabled = False
        chkInfoID3V2.Enabled = True
        chkInfoLyrics.Enabled = True
        chkInfoSyncLyrics.Enabled = True
        chkInfoPictures.Enabled = True
    End If
    
    lvwMP3_Click
End Sub

Private Sub chkID3V11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lvwMP3.SetFocus
End Sub

Private Sub chkID3V230_Click()
    If chkID3V230.Value = 1 Then
        chkID3V11.Value = 0
    Else
        If chkID3V11.Value = 0 Then chkID3V230.Value = 1
    End If
    
    If chkID3V11.Value = 1 Then
        cmdCopyToID3V11.Enabled = False
        cmdCopyToID3V230.Enabled = True
        chkInfoGeneral.Value = 1
    Else
        cmdCopyToID3V11.Enabled = True
        cmdCopyToID3V230.Enabled = False
    End If
    
    lvwMP3_Click
End Sub

Private Sub chkID3V230_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lvwMP3.SetFocus
End Sub

Private Sub chkInfoGeneral_Click()
    If chkInfoGeneral.Value = 1 Then
        chkInfoID3V2.Value = 0
        chkInfoLyrics.Value = 0
        chkInfoSyncLyrics.Value = 0
        chkInfoPictures.Value = 0
    Else
        If chkInfoID3V2.Value = 0 And chkInfoLyrics.Value = 0 And chkInfoSyncLyrics.Value = 0 And chkInfoPictures.Value = 0 Then chkInfoGeneral.Value = 1
    End If
    
    If chkInfoGeneral.Value = 1 Then
        fraGeneral.Visible = True
        fraID3V2.Visible = False
        fraLyrics.Visible = False
        fraSyncLyrics.Visible = False
        fraPictures.Visible = False
    End If
End Sub

Private Sub chkInfoGeneral_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lvwMP3.SetFocus
End Sub

Private Sub chkInfoID3V2_Click()
    If chkInfoID3V2.Value = 1 Then
        chkInfoGeneral.Value = 0
        chkInfoLyrics.Value = 0
        chkInfoSyncLyrics.Value = 0
        chkInfoPictures.Value = 0
    Else
        If chkInfoGeneral.Value = 0 And chkInfoLyrics.Value = 0 And chkInfoSyncLyrics.Value = 0 And chkInfoPictures.Value = 0 Then chkInfoID3V2.Value = 1
    End If
    
    If chkInfoID3V2.Value = 1 Then
        fraGeneral.Visible = False
        fraID3V2.Visible = True
        fraLyrics.Visible = False
        fraSyncLyrics.Visible = False
        fraPictures.Visible = False
    End If
End Sub

Private Sub chkInfoID3V2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lvwMP3.SetFocus
End Sub

Private Sub chkInfoLyrics_Click()
    If chkInfoLyrics.Value = 1 Then
        chkInfoID3V2.Value = 0
        chkInfoGeneral.Value = 0
        chkInfoSyncLyrics.Value = 0
        chkInfoPictures.Value = 0
    Else
        If chkInfoID3V2.Value = 0 And chkInfoSyncLyrics.Value = 0 And chkInfoGeneral.Value = 0 And chkInfoPictures.Value = 0 Then chkInfoLyrics.Value = 1
    End If
    
    If chkInfoLyrics.Value = 1 Then
        fraGeneral.Visible = False
        fraID3V2.Visible = False
        fraLyrics.Visible = True
        fraSyncLyrics.Visible = False
        fraPictures.Visible = False
    End If
End Sub

Private Sub chkInfoLyrics_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lvwMP3.SetFocus
End Sub

Private Sub chkInfoPictures_Click()
    If chkInfoPictures.Value = 1 Then
        chkInfoID3V2.Value = 0
        chkInfoLyrics.Value = 0
        chkInfoSyncLyrics.Value = 0
        chkInfoGeneral.Value = 0
    Else
        If chkInfoID3V2.Value = 0 And chkInfoSyncLyrics.Value = 0 And chkInfoLyrics.Value = 0 And chkInfoGeneral.Value = 0 Then chkInfoPictures.Value = 1
    End If
    
    If chkInfoPictures.Value = 1 Then
        fraGeneral.Visible = False
        fraID3V2.Visible = False
        fraLyrics.Visible = False
        fraSyncLyrics.Visible = False
        fraPictures.Visible = True
        cboPictureType_Click
    End If
End Sub

Private Sub chkInfoPictures_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lvwMP3.SetFocus
End Sub

Private Sub chkInfoSyncLyrics_Click()
    If chkInfoSyncLyrics.Value = 1 Then
        chkInfoID3V2.Value = 0
        chkInfoLyrics.Value = 0
        chkInfoPictures.Value = 0
        chkInfoGeneral.Value = 0
    Else
        If chkInfoID3V2.Value = 0 And chkInfoPictures.Value = 0 And chkInfoLyrics.Value = 0 And chkInfoGeneral.Value = 0 Then chkInfoSyncLyrics.Value = 1
    End If
    
    If chkInfoSyncLyrics.Value = 1 Then
        fraGeneral.Visible = False
        fraID3V2.Visible = False
        fraLyrics.Visible = False
        fraSyncLyrics.Visible = True
        fraPictures.Visible = False
    End If
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
    Unload frmAbout
End Sub

Private Sub cmdAddDirectory_Click()
    Dim intFile As Integer
    Dim strDir As String
    
    On Local Error GoTo ErrHan
    
    If cboDirectories.Text = "" Then Exit Sub
    strDir = cboDirectories.Text
    
    intFile = FreeFile
    Open AppendPath(App.Path) & "Directories.txt" For Append As intFile
        Write #intFile, cboDirectories.Text
    Close intFile
    
    LoadDirectories
    cboDirectories.Text = strDir
    
ErrHan:
    Close intFile
End Sub

Private Sub cmdAddURL_Click()
    Dim strURl As String
    
    strURl = InputBox("Please enter a URL that you wish to add:")
    If strURl = "" Then Exit Sub
    
    lstURL.AddItem strURl
End Sub

Private Sub cmdBatchOptions_Click()
    Dim lngCounter As Long
    Dim ID3V11 As New clsID3V111Writer
    Dim ID3V230 As New clsID3V230Writer
    Dim ID3V11Read As New clsID3V111Reader
    Dim ID3V230Read As New clsID3V230Reader
    Dim dblPercent As Double
    Dim strDirs() As String
    Dim strItems() As String
    Dim intNumber As Integer
    Dim strLyricsFile As String
    
    On Local Error GoTo Stopped
    
    typExample.strDir = lvwMP3.ListItems(1).SubItems(7)
    typExample.strFile = lvwMP3.ListItems(1).Text
    If typExample.strDir = "" Then Exit Sub
    
    frmBatchOptions.Show vbModal, Me
    Unload frmBatchOptions
    DoEvents
    
    If typBatchOptions.bolCancelled = True Then Exit Sub
    
    fbolStopBatch = False
    DisableControls
    For lngCounter = 1 To lvwMP3.ListItems.Count
        intNumber = 0
        dblPercent = (lngCounter / lvwMP3.ListItems.Count) * 100
        
        strDirs = Split(lvwMP3.ListItems(lngCounter).SubItems(7), "\")
        strItems = Split(Left(lvwMP3.ListItems(lngCounter).Text, Len(lvwMP3.ListItems(lngCounter).Text) - 4), typBatchOptions.strSeparator)
        
        If typBatchOptions.bolRemoveID3V111 = True Then
            barStatus.Panels(1).Text = "Removing ID3V1 1.1 tag from " & lvwMP3.ListItems(lngCounter).Text & " (" & Format(dblPercent, "0.00") & "%)" & " [Click here to cancel]"
            DoEvents
            ID3V11.RemoveID3V111Tag lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text
        End If
        
        If typBatchOptions.bolRemoveID3V230 = True Then
            barStatus.Panels(1).Text = "Removing ID3V2 3.0 tag from " & lvwMP3.ListItems(lngCounter).Text & " (" & Format(dblPercent, "0.00") & "%)" & " [Click here to cancel]"
            DoEvents
            ID3V230.RemoveID3V230Tag lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text
        End If
        
        If typBatchOptions.bolCopyID3V111ToID3V230 Then
            'Read in V230 Basic Info, Copy it to V111
            ID3V11Read.ReadID3V111Tag lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text
            ID3V230.Artist = ID3V11Read.Artist
            ID3V230.Album = ID3V11Read.Album
            ID3V230.SongTitle = ID3V11Read.SongTitle
            ID3V230.Track = ID3V11Read.Track
            ID3V230.Year = ID3V11Read.Year
            ID3V230.Genre = ID3V11Read.Genre
            ID3V230.Comment = ID3V11Read.Comment
            barStatus.Panels(1).Text = "Copying ID3V1 V1.1 tag from " & lvwMP3.ListItems(lngCounter).Text & " (" & Format(dblPercent, "0.00") & "%)" & " [Click here to cancel]"
            DoEvents
            ID3V230.WriteID3V230Tag lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text
        End If
        
        If typBatchOptions.bolCopyID3V230ToID3V111 Then
            'Read in V111 Info, Copy it to V230
            ID3V230Read.ReadID3V230Tag lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text, True
            ID3V11.Artist = ID3V230Read.Artist
            ID3V11.Album = ID3V230Read.Album
            ID3V11.SongTitle = ID3V230Read.SongTitle
            
            If InStr(ID3V230Read.Track, "/") = 0 Then
                ID3V11.Track = ID3V230Read.Track
            Else
                ID3V11.Track = Left(ID3V230Read.Track, InStr(ID3V230Read.Track, "/") - 1)
            End If
            
            ID3V11.Year = ID3V230Read.Year
            ID3V11.Genre = ID3V11.GetGenre(ID3V230Read.Genre)
            ID3V11.Comment = ID3V230Read.Comment
            barStatus.Panels(1).Text = "Copying ID3V2 V3.0 tag from " & lvwMP3.ListItems(lngCounter).Text & " (" & Format(dblPercent, "0.00") & "%)" & " [Click here to cancel]"
            DoEvents
            ID3V11.WriteID3V111Tag lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text
        End If
        
        If typBatchOptions.bolRewriteID3V111 = True Then
            With ID3V11
                On Local Error Resume Next
                
                If Len(typBatchOptions.strArtist) <> 0 Then
                    Select Case LCase(typBatchOptions.strArtist)
                        Case "d"
                            .Artist = strDirs(UBound(strDirs) - 1)
                        Case "u"
                            .Artist = strDirs(UBound(strDirs) - 2)
                        Case Else
                            If IsNumeric(typBatchOptions.strArtist) = True Then
                                intNumber = CInt(typBatchOptions.strArtist)
                                .Artist = strItems(intNumber - 1)
                            End If
                    End Select
                    
                    If typBatchOptions.strArtistOverride <> "" Then .Artist = typBatchOptions.strArtistOverride
                End If
                
                If Len(typBatchOptions.strAlbum) <> 0 Then
                    Select Case LCase(typBatchOptions.strAlbum)
                        Case "d"
                            .Album = strDirs(UBound(strDirs) - 1)
                        Case "u"
                            .Album = strDirs(UBound(strDirs) - 2)
                        Case Else
                            If IsNumeric(typBatchOptions.strAlbum) = True Then
                                intNumber = CInt(typBatchOptions.strAlbum)
                                .Album = strItems(intNumber - 1)
                            End If
                    End Select
                    
                    If typBatchOptions.strAlbumOverride <> "" Then .Album = typBatchOptions.strAlbumOverride
                End If
                
                If Len(typBatchOptions.strTrack) <> 0 Then
                    Select Case LCase(typBatchOptions.strTrack)
                        Case "d"
                            .Track = strDirs(UBound(strDirs) - 1)
                        Case "u"
                            .Track = strDirs(UBound(strDirs) - 2)
                        Case Else
                            If IsNumeric(typBatchOptions.strTrack) = True Then
                                intNumber = CInt(typBatchOptions.strTrack)
                                .Track = strItems(intNumber - 1)
                            End If
                    End Select
                    
                    If typBatchOptions.intTrackOverride <> -1 Then .Track = typBatchOptions.intTrackOverride
                End If
                
                If Len(typBatchOptions.strSong) <> 0 Then
                    Select Case LCase(typBatchOptions.strSong)
                        Case "d"
                            .SongTitle = strDirs(UBound(strDirs) - 1)
                        Case "u"
                            .SongTitle = strDirs(UBound(strDirs) - 2)
                        Case Else
                            If IsNumeric(typBatchOptions.strSong) = True Then
                                intNumber = CInt(typBatchOptions.strSong)
                                .SongTitle = strItems(intNumber - 1)
                            End If
                    End Select
                    
                    If typBatchOptions.strSongOverride <> "" Then .SongTitle = typBatchOptions.strSongOverride
                End If
                
                .Comment = typBatchOptions.strCommentOverride
                If typBatchOptions.strGenreOverride <> "" Then .Genre = .GetGenre(typBatchOptions.strGenreOverride)
                .Year = typBatchOptions.strYearOverride
                
                barStatus.Panels(1).Text = "Writing ID3V1 1.1 tag to " & lvwMP3.ListItems(lngCounter).Text & " (" & Format(dblPercent, "0.00") & "%)" & " [Click here to cancel]"
                DoEvents
                .WriteID3V111Tag lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text
            End With
        End If
        
        If typBatchOptions.bolRewriteID3V230 = True Then
            With ID3V230
                On Local Error Resume Next
                
                If Len(typBatchOptions.strArtist) <> 0 Then
                    Select Case LCase(typBatchOptions.strArtist)
                        Case "d"
                            .Artist = strDirs(UBound(strDirs) - 1)
                        Case "u"
                            .Artist = strDirs(UBound(strDirs) - 2)
                        Case Else
                            If IsNumeric(typBatchOptions.strArtist) = True Then
                                intNumber = CInt(typBatchOptions.strArtist)
                                .Artist = strItems(intNumber - 1)
                            End If
                    End Select
                    
                    If typBatchOptions.strArtistOverride <> "" Then .Artist = typBatchOptions.strArtistOverride
                End If
                
                If Len(typBatchOptions.strAlbum) <> 0 Then
                    Select Case LCase(typBatchOptions.strAlbum)
                        Case "d"
                            .Album = strDirs(UBound(strDirs) - 1)
                        Case "u"
                            .Album = strDirs(UBound(strDirs) - 2)
                        Case Else
                            If IsNumeric(typBatchOptions.strAlbum) = True Then
                                intNumber = CInt(typBatchOptions.strAlbum)
                                .Album = strItems(intNumber - 1)
                            End If
                    End Select
                    
                    If typBatchOptions.strAlbumOverride <> "" Then .Album = typBatchOptions.strAlbumOverride
                End If
                
                If Len(typBatchOptions.strTrack) <> 0 Then
                    Select Case LCase(typBatchOptions.strTrack)
                        Case "d"
                            .Track = strDirs(UBound(strDirs) - 1)
                        Case "u"
                            .Track = strDirs(UBound(strDirs) - 2)
                        Case Else
                            If IsNumeric(typBatchOptions.strTrack) = True Then
                                intNumber = CInt(typBatchOptions.strTrack)
                                .Track = strItems(intNumber - 1)
                            End If
                    End Select
                    
                    If typBatchOptions.intTrackOverride <> -1 Then .Track = typBatchOptions.intTrackOverride
                End If
                
                If Len(typBatchOptions.strSong) <> 0 Then
                    Select Case LCase(typBatchOptions.strSong)
                        Case "d"
                            .SongTitle = strDirs(UBound(strDirs) - 1)
                        Case "u"
                            .SongTitle = strDirs(UBound(strDirs) - 2)
                        Case Else
                            If IsNumeric(typBatchOptions.strSong) = True Then
                                intNumber = CInt(typBatchOptions.strSong)
                                .SongTitle = strItems(intNumber - 1)
                            End If
                    End Select
                    
                    If typBatchOptions.strSongOverride <> "" Then .SongTitle = typBatchOptions.strSongOverride
                End If
                
                .Comment = typBatchOptions.strCommentOverride
                If typBatchOptions.strGenreOverride <> "" Then .Genre = typBatchOptions.strGenreOverride
                .Year = typBatchOptions.strYearOverride
                
                strLyricsFile = ""
                .Lyrics = ""
                If typBatchOptions.bolWriteLyrics = True Then
                    strLyricsFile = Left(lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text, Len(lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text) - 4) & ".txt"
                    If FileExists(strLyricsFile) = True Then
                        .Lyrics = ReadFile(strLyricsFile)
                    End If
                End If
                
                .Graphic = ""
                .GraphicTitle = ""
                .GraphicType = ""
                If typBatchOptions.bolWritePictures = True Then
                    If FileExists(lvwMP3.ListItems(lngCounter).SubItems(7) & "Cover.jpg") = True Then
                        .Graphic = lvwMP3.ListItems(lngCounter).SubItems(7) & "Cover.jpg"
                        .GraphicTitle = "Album Front Cover"
                        .GraphicType = .GetGraphicType("Front Cover")
                    End If
                    
                    If FileExists(lvwMP3.ListItems(lngCounter).SubItems(7) & "Back.jpg") = True Then
                        If .Graphic = "" Then
                            .Graphic = lvwMP3.ListItems(lngCounter).SubItems(7) & "Back.jpg"
                            .GraphicTitle = "Album Back Cover"
                            .GraphicType = .GetGraphicType("Back Cover")
                        Else
                            .Graphic = .Graphic & "|" & lvwMP3.ListItems(lngCounter).SubItems(7) & "Back.jpg"
                            .GraphicTitle = .GraphicTitle & "|" & "Album Back Cover"
                            .GraphicType = .GetGraphicType(.GraphicType & "|" & "Back Cover")
                        End If
                    End If
                    
                    If FileExists(lvwMP3.ListItems(lngCounter).SubItems(7) & "Band.jpg") = True Then
                        If .Graphic = "" Then
                            .Graphic = lvwMP3.ListItems(lngCounter).SubItems(7) & "Band.jpg"
                            .GraphicTitle = "Band Picture"
                            .GraphicType = .GetGraphicType("Band")
                        Else
                            .Graphic = .Graphic & "|" & lvwMP3.ListItems(lngCounter).SubItems(7) & "Band.jpg"
                            .GraphicTitle = .GraphicTitle & "|" & "Band Picture"
                            .GraphicType = .GetGraphicType(.GraphicType & "|" & "Band")
                        End If
                    End If
                End If
                
                barStatus.Panels(1).Text = "Writing ID3V2 3.0 tag to " & lvwMP3.ListItems(lngCounter).Text & " (" & Format(dblPercent, "0.00") & "%)" & " [Click here to cancel]"
                DoEvents
                .WriteID3V230Tag lvwMP3.ListItems(lngCounter).SubItems(7) & lvwMP3.ListItems(lngCounter).Text
            End With
        End If
        
        If fbolStopBatch = True Then
            fbolStopBatch = False
            GoTo Stopped
        End If
    Next
    
    cmdScanDirectory_Click
    
Stopped:
    EnableControls
    barStatus.Panels(1).Text = "Ready."
End Sub

Private Sub cmdCopyFromLyrics_Click()
    If txtLyrics.Text <> "" Then txtSyncLyrics.Text = txtLyrics.Text
End Sub

Private Sub cmdCopyFromSync_Click()
    Dim strLines() As String
    Dim lngCounter As Long
    Dim strLine As String
    Dim strLyrics As String
    
    On Local Error Resume Next
    
    If txtSyncLyrics.Text <> "" Then
        strLines = Split(txtSyncLyrics.Text, vbCrLf)
        For lngCounter = 0 To UBound(strLines)
            strLine = Right(strLines(lngCounter), Len(strLines(lngCounter)) - InStr(strLines(lngCounter), "]"))
            If strLyrics = "" Then
                strLyrics = strLine
            Else
                strLyrics = strLyrics & vbCrLf & strLine
            End If
        Next
        txtLyrics.Text = strLyrics
    End If
End Sub

Private Sub cmdCopyToID3V11_Click()
    Dim ID3V11 As New clsID3V111Writer

    On Local Error Resume Next

    With ID3V11
        .Artist = txtArtist.Text
        .Album = txtAlbum.Text
        .SongTitle = txtSong.Text
        .Track = IIf(IsNumeric(txtTrack.Text) = True, CInt(txtTrack.Text), 0)
        .Year = txtYear.Text
        .Genre = IIf(cboGenre.Text <> "", .GetGenre(cboGenre.Text), 255)
        .Comment = txtComment.Text
        fraID3.Enabled = False
        .WriteID3V111Tag lvwMP3.SelectedItem.SubItems(7) & lvwMP3.SelectedItem
        fraID3.Enabled = True
        chkID3V11.Value = 1
    End With
End Sub

Private Sub cmdCopyToID3V230_Click()
    Dim ID3V230 As New clsID3V230Writer

    On Local Error Resume Next

    With ID3V230
        .Artist = txtArtist.Text
        .Album = txtAlbum.Text
        .SongTitle = txtSong.Text
        .Track = txtTrack.Text
        .Year = txtYear.Text
        .Genre = cboGenre.Text
        .Comment = txtComment.Text
        fraID3.Enabled = False
        .WriteID3V230Tag lvwMP3.SelectedItem.SubItems(7) & lvwMP3.SelectedItem
        fraID3.Enabled = True
        chkID3V230.Value = 1
    End With
End Sub

Private Sub cmdDeleteURL_Click()
    If lstURL.ListIndex = -1 Then Exit Sub
    
    If MsgBox("Are you sure you want to delete this URL?", vbYesNo + vbQuestion) = vbYes Then lstURL.RemoveItem lstURL.ListIndex
End Sub

Private Sub cmdInsertTime_Click()
    Dim strTime As String
    
    strTime = lblTime.Caption
    
    txtSyncLyrics.SetFocus
    SendKeys "[" & strTime & "]"
    SendKeys "{DOWN}"
    SendKeys "{HOME}"
End Sub

Private Sub cmdLoadSyncLyrics_Click()
    Dim strFileName As String
    Dim intFile As Integer
    Dim strLine As String
    
    On Local Error GoTo ErrHan
    
    dlgOpen.FileName = ""
    dlgOpen.Filter = "Text File|*.txt|Syncronized Lyrics File|*.lrc"
    
    dlgOpen.ShowOpen
    strFileName = dlgOpen.FileName
    
    If FileExists(strFileName) = False Then GoTo ErrHan
    
    intFile = FreeFile
    txtSyncLyrics.Text = ""
    If LCase(Left(strFileName, 3)) = "txt" Then
        Open strFileName For Input As intFile
            While Not EOF(intFile)
                Line Input #intFile, strLine
                If txtSyncLyrics.Text = "" Then
                    txtSyncLyrics.Text = strLine
                Else
                    txtSyncLyrics.Text = txtSyncLyrics.Text & vbCrLf & strLine
                End If
            Wend
        Close intFile
    Else
        ReDim Lyrics(0)
        intFile = FreeFile
        Open strFileName For Input As intFile
            While Not EOF(intFile)
                Line Input #intFile, strLine
                ProcessLine strLine
            Wend
        Close intFile
        
        SortLyrics
        DisplayLyrics
        ReDim Lyrics(0)
    End If
    
    Exit Sub
    
ErrHan:
    MsgBox Err.Description
End Sub

Private Sub cmdSaveLyrics_Click()
    Dim strFileName As String
    Dim intFile As Integer
    
    On Local Error GoTo ErrHan
    
    If txtLyrics.Text = "" Then Exit Sub
    dlgSave.FileName = ""
    dlgSave.Filter = "Text Files|*.txt"
    
    dlgSave.ShowSave
    strFileName = dlgSave.FileName
    
    intFile = FreeFile
    Open strFileName For Output As intFile
        Print #intFile, txtLyrics.Text
    Close intFile
    
    Exit Sub
    
ErrHan:
End Sub

Private Sub cmdEditURL_Click()
    Dim strURl As String
    
    If lstURL.ListIndex = -1 Then Exit Sub
    
    strURl = InputBox("Change URL to:", , lstURL.List(lstURL.ListIndex))
    
    If strURl <> "" Then lstURL.List(lstURL.ListIndex) = strURl
End Sub

Private Sub cmdLoadLyrics_Click()
    Dim strFileName As String
    Dim intFile As Integer
    Dim strLine As String
    
    On Local Error GoTo ErrHan
    
    dlgOpen.FileName = ""
    dlgOpen.Filter = "Text Files|*.txt"
    
    dlgOpen.ShowOpen
    strFileName = dlgOpen.FileName
    
    If FileExists(strFileName) = False Then GoTo ErrHan
    
    intFile = FreeFile
    txtLyrics.Text = ""
    Open strFileName For Input As intFile
        While Not EOF(intFile)
            Line Input #intFile, strLine
            If txtLyrics.Text = "" Then
                txtLyrics.Text = strLine
            Else
                txtLyrics.Text = txtLyrics.Text & vbCrLf & strLine
            End If
        Wend
    Close intFile
    
    Exit Sub
    
ErrHan:
End Sub

Private Sub cmdLoadPicture_Click()
    Dim lngCounter As Long
    Dim strFileName As String
    
    On Local Error GoTo ErrHan
    
    If cboPictureType.Text = "" Then Exit Sub
    
    dlgOpen.FileName = ""
    dlgOpen.Filter = "JPEG Files|*.jpg|Gif Files|*.gif|PNG Files|*.png|Bitmap Files|*.bmp|All Files|*.*"
    
    dlgOpen.ShowOpen
    strFileName = dlgOpen.FileName
    
    picPicture = LoadPicture(strFileName)
    
    For lngCounter = 1 To UBound(fGraphics)
        If fGraphics(lngCounter).bytType = cboPictureType.ItemData(cboPictureType.ListIndex) Then
            fGraphics(lngCounter).strFileName = strFileName
            Exit For
        End If
    Next
    
    If lngCounter = UBound(fGraphics) + 1 Then
        ReDim Preserve fGraphics(UBound(fGraphics) + 1)
        fGraphics(UBound(fGraphics)).strFileName = strFileName
        fGraphics(UBound(fGraphics)).bytType = cboPictureType.ItemData(cboPictureType.ListIndex)
    End If
    
    Exit Sub
    
ErrHan:
End Sub

Private Sub cmdRemoveDirectory_Click()
    Dim strDir As String
    Dim lngCounter As Long
    Dim intFile As Integer
    
    On Local Error GoTo ErrHan
    
    If cboDirectories.Text = "" Then Exit Sub
    
    strDir = cboDirectories.Text
    intFile = FreeFile
    Open AppendPath(App.Path) & "Directories.txt" For Output As intFile
        For lngCounter = 0 To cboDirectories.ListCount - 1
            If cboDirectories.List(lngCounter) <> strDir Then Write #intFile, cboDirectories.List(lngCounter)
        Next
    Close intFile
    
    LoadDirectories
    
ErrHan:
    Close intFile
End Sub

Private Sub cmdRemoveID3_Click()
    Dim ID3V11 As New clsID3V111Writer
    Dim ID3V230 As New clsID3V230Writer
    
    If chkID3V11.Value = 1 Then
        fraID3.Enabled = False
        ID3V11.RemoveID3V111Tag lvwMP3.SelectedItem.SubItems(7) & lvwMP3.SelectedItem
        fraID3.Enabled = True
        lvwMP3_Click
    Else
        fraID3.Enabled = False
        ID3V230.RemoveID3V230Tag lvwMP3.SelectedItem.SubItems(7) & lvwMP3.SelectedItem
        fraID3.Enabled = True
        lvwMP3_Click
    End If
End Sub

Private Sub cmdRemovePicture_Click()
    Dim lngCounter As Long
    
    On Local Error GoTo ErrHan
    
    If (picPicture.Picture Is Nothing) Or (cboPictureType.Text = "") Then Exit Sub
    
    For lngCounter = 1 To UBound(fGraphics)
        If fGraphics(lngCounter).bytType = cboPictureType.ItemData(cboPictureType.ListIndex) Then
            fGraphics(lngCounter).strFileName = ""
            Exit For
        End If
    Next
    
    Exit Sub
    
ErrHan:
End Sub

Private Sub cmdSaveID3_Click()
    Dim ID3V11 As New clsID3V111Writer
    Dim ID3V230 As New clsID3V230Writer
    Dim lngCounter As Long
    Dim strURl As String
    Dim strGraphic As String
    Dim strGraphicType As String
    Dim strGraphicTitle As String
    
    If chkID3V11.Value = 1 Then
        With ID3V11
            .Artist = txtArtist.Text
            .Album = txtAlbum.Text
            .SongTitle = txtSong.Text
            If txtTrack.Text <> "" Then
                .Track = IIf(IsNumeric(txtTrack.Text) = True, CInt(txtTrack.Text), 0)
            End If
            .Year = txtYear.Text
            .Genre = IIf(cboGenre.Text <> "", .GetGenre(cboGenre.Text), 255)
            .Comment = txtComment.Text
            fraID3.Enabled = False
            .WriteID3V111Tag lvwMP3.SelectedItem.SubItems(7) & lvwMP3.SelectedItem
            fraID3.Enabled = True
        End With
        
        lvwMP3_Click
    Else
        With ID3V230
            .Artist = txtArtist.Text
            .Album = txtAlbum.Text
            .SongTitle = txtSong.Text
            .Track = txtTrack.Text
            .Year = txtYear.Text
            .Genre = cboGenre.Text
            .Comment = txtComment.Text
            .OriginalArtist = txtOriginalArtist.Text
            .SubTitle = txtSubtitle.Text
            .Composer = txtComposer.Text
            .Copyright = txtCopyright.Text
            .EncodedBy = txtEncodedBy.Text
            .Language = txtLanguage.Text
            
            For lngCounter = 0 To lstURL.ListCount - 1
                If strURl = "" Then
                    strURl = lstURL.List(lngCounter)
                Else
                    strURl = strURl & "|" & lstURL.List(lngCounter)
                End If
            Next
            
            .Lyrics = txtLyrics.Text
            .SyncLyrics = txtSyncLyrics.Text
            
            For lngCounter = 1 To UBound(fGraphics)
                If FileExists(fGraphics(lngCounter).strFileName) = True Then
                    If strGraphic = "" Then
                        strGraphic = fGraphics(lngCounter).strFileName
                        strGraphicType = fGraphics(lngCounter).bytType
                        strGraphicTitle = fGraphics(lngCounter).strTitle
                    Else
                        strGraphic = strGraphic & "|" & fGraphics(lngCounter).strFileName
                        strGraphicType = strGraphicType & "|" & fGraphics(lngCounter).bytType
                        strGraphicTitle = strGraphicTitle & "|" & fGraphics(lngCounter).strTitle
                    End If
                End If
            Next
            
            .Graphic = strGraphic
            .GraphicType = strGraphicType
            .GraphicTitle = strGraphicTitle
            
            fraID3.Enabled = False
            Me.MousePointer = vbHourglass
            .WriteID3V230Tag lvwMP3.SelectedItem.SubItems(7) & lvwMP3.SelectedItem
            Me.MousePointer = vbNormal
            fraID3.Enabled = True
        End With
            
        lvwMP3_Click
    End If
End Sub

Private Sub cmdSavePicture_Click()
    Dim lngCounter As Long
    Dim strFileName As String
    Dim Reader As New clsID3V230Reader
    Dim strTempFile As String
    
    On Local Error GoTo ErrHan
    
    If (picPicture.Picture Is Nothing) Or (cboPictureType.Text = "") Then Exit Sub
    dlgSave.FileName = ""
    dlgSave.Filter = "All Files|*.*"
    
    dlgSave.ShowSave
    strFileName = dlgSave.FileName
    
    For lngCounter = 1 To UBound(fGraphics)
        If fGraphics(lngCounter).bytType = cboPictureType.ItemData(cboPictureType.ListIndex) Then
            Me.MousePointer = vbHourglass
            strTempFile = Reader.ExtractGraphic(lvwMP3.SelectedItem.SubItems(7) & lvwMP3.SelectedItem, fGraphics(lngCounter).dblStartPosition, fGraphics(lngCounter).dblLength)
            Name strTempFile As strFileName
            Me.MousePointer = vbNormal
            Exit For
        End If
    Next
    
    Exit Sub
    
ErrHan:
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdSaveSyncLyrics_Click()
    Dim strFileName As String
    Dim intFile As Integer
    
    On Local Error GoTo ErrHan
    
    If txtSyncLyrics.Text = "" Then Exit Sub
    dlgSave.FileName = ""
    dlgSave.Filter = "Text Files|*.txt"
    
    dlgSave.ShowSave
    strFileName = dlgSave.FileName
    
    intFile = FreeFile
    Open strFileName For Output As intFile
        Print #intFile, txtSyncLyrics.Text
    Close intFile
    
    Exit Sub
    
ErrHan:
End Sub

Private Sub cmdScanDirectory_Click()
    Select Case cmdScanDirectory.Caption
        Case "&Scan"
            ClearControls
            lvwMP3.ListItems.Clear
            lvwMP3.Enabled = False
            LockWindowUpdate lvwMP3.hwnd
            fbolStop = False
            cmdScanDirectory.Caption = "&Stop"
            barStatus.Panels(1).Text = "Scanning For Files..."
            DoSearch cboDirectories.Text, "*.mp3"
            barStatus.Panels(1).Text = "Ready."
            cmdScanDirectory.Caption = "&Scan"
            lvwMP3.Enabled = True
            LockWindowUpdate 0
            DoEvents
        Case Else
            cmdScanDirectory.Caption = "&Scan"
            fbolStop = True
    End Select
End Sub

Private Sub cmdSetTitle_Click()
    Dim lngCounter As Long
    
    If cboPictureType.ListIndex = -1 Then Exit Sub
        
    For lngCounter = 1 To UBound(fGraphics)
        If fGraphics(lngCounter).bytType = cboPictureType.ItemData(cboPictureType.ListIndex) Then
            fGraphics(lngCounter).strTitle = txtPictureTitle.Text
        End If
    Next
End Sub

Private Sub Form_Load()
    barStatus.Panels(1).Text = "Ready."
    
    lvwMP3.ColumnHeaders.Clear
    lvwMP3.ColumnHeaders.Add , , "File", (lvwMP3.Width / 3) * 1.5
    lvwMP3.ColumnHeaders.Add , , "MPEG", (2 * (lvwMP3.Width / 3)) * 0.15
    lvwMP3.ColumnHeaders.Add , , "Bitrate/Frequency", (2 * (lvwMP3.Width / 3)) * 0.15
    lvwMP3.ColumnHeaders.Add , , "Mode", (2 * (lvwMP3.Width / 3)) * 0.1
    lvwMP3.ColumnHeaders.Add , , "Length", (2 * (lvwMP3.Width / 3)) * 0.1
    lvwMP3.ColumnHeaders.Add , , "ID3V1 V1.1", (2 * (lvwMP3.Width / 3)) * 0.1
    lvwMP3.ColumnHeaders.Add , , "ID3V2 V3.0", (2 * (lvwMP3.Width / 3)) * 0.1
    lvwMP3.ColumnHeaders.Add , , "", 0
    lvwMP3.View = lvwReport
    chkID3V230_Click
    chkInfoGeneral_Click
    ReDim fGraphics(0)
    
    LoadDirectories
    LoadGenres
    LoadPictureTypes
    
    On Local Error Resume Next
    cboDirectories.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    
    lvwMP3.Width = Me.Width - 345
    lvwMP3.Height = Me.Height - 6375
    cboDirectories.Width = Me.Width - 345
    fraFrame.Left = Me.Width - 3720
    fraID3Buttons.Width = Me.Width - 345
    fraID3Buttons.Top = Me.Height - 5070
    fraInfo.Width = Me.Width - 345
    fraInfo.Top = Me.Height - 4590
    fraID3.Width = Me.Width - 345
    fraID3.Top = Me.Height - 1350
    cmdAbout.Left = fraID3.Width - cmdAbout.Width
    
    lvwMP3.ColumnHeaders(1).Width = (lvwMP3.Width / 3) * 1.5
    lvwMP3.ColumnHeaders(2).Width = (2 * (lvwMP3.Width / 3)) * 0.15
    lvwMP3.ColumnHeaders(3).Width = (2 * (lvwMP3.Width / 3)) * 0.15
    lvwMP3.ColumnHeaders(4).Width = (2 * (lvwMP3.Width / 3)) * 0.1
    lvwMP3.ColumnHeaders(5).Width = (2 * (lvwMP3.Width / 3)) * 0.1
    lvwMP3.ColumnHeaders(6).Width = (2 * (lvwMP3.Width / 3)) * 0.1
    lvwMP3.ColumnHeaders(7).Width = (2 * (lvwMP3.Width / 3)) * 0.1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngCounter As Long
    
    On Local Error Resume Next
    
    For lngCounter = 1 To UBound(fGraphics)
        If fGraphics(lngCounter).bolExtracted = True Then Kill fGraphics(lngCounter).strFileName
    Next
    
    fbolStop = True
End Sub

Private Sub DoSearch(ByVal strPath As String, ByVal strPattern As String)
    Dim strFile As String
    Dim lngCounter As Long
    Dim strFolders() As String
    Dim MP3Info As New clsMP3HeadReader
    Dim ID3V111 As New clsID3V111Reader
    Dim ID3V230 As New clsID3V230Reader
    Dim MP3Time As MS
    
    On Local Error Resume Next
    
    'Get the files
    strPath = AppendPath(strPath)
    barStatus.Panels(1).Text = "Searching..." & strPath
    strFile = Dir(strPath & strPattern, vbNormal + vbHidden + vbSystem + vbReadOnly + vbArchive + vbDirectory)
    Do Until (strFile = "") Or (fbolStop = True)
        If (GetAttr(strPath & strFile) And vbDirectory) <> vbDirectory Then
            MP3Info.ReadMP3Header strPath & strFile
            lvwMP3.ListItems.Add , , strFile
            lvwMP3.ListItems(lvwMP3.ListItems.Count).SubItems(1) = MP3Info.MPEG & " Layer " & MP3Info.Layer
            lvwMP3.ListItems(lvwMP3.ListItems.Count).SubItems(2) = MP3Info.Bitrate & "/" & MP3Info.Frequency
            lvwMP3.ListItems(lvwMP3.ListItems.Count).SubItems(3) = MP3Info.Mode
            MP3Time = TMS(MP3Info.Length)
            lvwMP3.ListItems(lvwMP3.ListItems.Count).SubItems(4) = MP3Time.M & ":" & Format(MP3Time.S, "00")
            ID3V111.ReadID3V111Tag strPath & strFile, True
            lvwMP3.ListItems(lvwMP3.ListItems.Count).SubItems(5) = ID3V111.TagPresent
            ID3V230.ReadID3V230Tag strPath & strFile, True, True
            lvwMP3.ListItems(lvwMP3.ListItems.Count).SubItems(6) = ID3V230.TagPresent
            lvwMP3.ListItems(lvwMP3.ListItems.Count).SubItems(7) = strPath
        End If
        strFile = Dir
        DoEvents
    Loop
    
    If fbolStop = True Then
        barStatus.Panels(1).Text = "Ready."
        lvwMP3.Enabled = True
        LockWindowUpdate 0
        DoEvents
        Exit Sub
    End If
    
    If chkRecursiveSearch.Value = 0 Then Exit Sub
    
    'Get the folders
    ReDim strFolders(0 To 0)
    strFile = Dir(strPath, vbNormal + vbHidden + vbSystem + vbReadOnly + vbArchive + vbDirectory)
    Do Until (strFile = "") Or (fbolStop = True)
        If (GetAttr(strPath & strFile) And vbDirectory) = vbDirectory Then
            If (strFile <> ".") And (strFile <> "..") Then
                ReDim Preserve strFolders(0 To UBound(strFolders) + 1)
                strFolders(UBound(strFolders)) = strPath & AppendPath(strFile)
            End If
        End If
        strFile = Dir
        DoEvents
    Loop
    
    If fbolStop = True Then
        barStatus.Panels(1).Text = "Ready."
        lvwMP3.Enabled = True
        LockWindowUpdate 0
        DoEvents
        Exit Sub
    End If
    
    'Recurse search
    For lngCounter = 1 To UBound(strFolders)
        If fbolStop = True Then Exit Sub
        DoSearch strFolders(lngCounter), strPattern
        DoEvents
    Next lngCounter
End Sub

Private Sub LoadDirectories()
    Dim intFile As Integer
    Dim strDir As String
    
    On Local Error GoTo ErrHan
    
    cboDirectories.Clear
    cboDirectories.Text = ""
    intFile = FreeFile
    Open AppendPath(App.Path) & "Directories.txt" For Input As intFile
        While Not EOF(intFile)
            Input #intFile, strDir
            cboDirectories.AddItem strDir
            DoEvents
        Wend
    Close intFile
    
ErrHan:
    Close intFile
End Sub

Private Sub lvwMP3_Click()
    On Local Error GoTo ErrHan
    
    ClearControls
    If lvwMP3.SelectedItem = "" Then Exit Sub
    LoadID3Info lvwMP3.SelectedItem.SubItems(7) & lvwMP3.SelectedItem
    
ErrHan:
End Sub

Private Sub LoadID3Info(ByVal strFile As String)
    Dim ID3V11 As New clsID3V111Reader
    Dim ID3V230 As New clsID3V230Reader
    Dim ID3V230X As New clsID3V230Writer
    Dim strURl() As String
    Dim strGraphicType() As String
    Dim strGraphicLength() As String
    Dim strGraphicStartPos() As String
    Dim strGraphicTitle() As String
    Dim lngCounter As Long
    Dim strFileName As String
    Dim strFilePart As String
    
    On Local Error GoTo ErrHan
    
    ID3V11.ReadID3V111Tag strFile
    lvwMP3.SelectedItem.SubItems(5) = ID3V11.TagPresent
    ID3V230.ReadID3V230Tag strFile
    lvwMP3.SelectedItem.SubItems(6) = ID3V230.TagPresent
    
    For lngCounter = 1 To UBound(fGraphics)
        If fGraphics(lngCounter).bolExtracted = True Then
            If FileExists(fGraphics(lngCounter).strFileName) = True Then
                Kill fGraphics(lngCounter).strFileName
                DoEvents
            End If
        End If
    Next
    
    ReDim fGraphics(0)
    
    If chkID3V11.Value = 1 Then
        With ID3V11
            If .TagPresent = True Then
                txtArtist.Text = .Artist
                txtAlbum.Text = .Album
                txtSong.Text = .SongTitle
                txtTrack.Text = .Track
                txtYear.Text = .Year
                txtComment.Text = .Comment
                cboGenre.Text = .Genre
            End If
        End With
    Else
        With ID3V230
            If .TagPresent = True Then
                txtArtist.Text = .Artist
                txtAlbum.Text = .Album
                txtSong.Text = .SongTitle
                txtTrack.Text = .Track
                txtYear.Text = .Year
                txtComment.Text = .Comment
                cboGenre.Text = .Genre
                txtOriginalArtist.Text = .OriginalArtist
                txtSubtitle.Text = .SubTitle
                txtComposer.Text = .Composer
                txtCopyright.Text = .Copyright
                txtEncodedBy.Text = .EncodedBy
                txtLanguage.Text = .Language
                
                If .URL <> "" Then
                    If InStr(.URL, "|") = 0 Then
                        lstURL.AddItem .URL
                    Else
                        strURl = Split(.URL, "|")
                        For lngCounter = 0 To UBound(strURl)
                            lstURL.AddItem strURl(lngCounter)
                        Next
                    End If
                End If
                
                txtUnknown.Text = Replace(.Unknown, "|", vbCrLf)
                txtLyrics.Text = .Lyrics
                If .SyncLyrics <> "" Then
                    txtSyncLyrics.Text = .SyncLyrics
                End If
                
                If .GraphicSize <> "" Then
                    If InStr(.GraphicSize, "|") = 0 Then
                        ReDim Preserve fGraphics(UBound(fGraphics) + 1)
                        fGraphics(UBound(fGraphics)).bytType = ID3V230X.GetGraphicType(.GraphicExtended)
                        fGraphics(UBound(fGraphics)).dblLength = .GraphicSize
                        fGraphics(UBound(fGraphics)).dblStartPosition = .GraphicStartPos
                        fGraphics(UBound(fGraphics)).strTitle = .GraphicTitle
                        strFileName = .ExtractGraphic(strFile, .GraphicStartPos, .GraphicSize)
                        strFilePart = GetFileName(strFileName)
                        If FileExists(AppendPath(App.Path) & strFilePart) = True Then Kill AppendPath(App.Path) & strFilePart
                        Name strFileName As AppendPath(App.Path) & strFilePart
                        fGraphics(UBound(fGraphics)).strFileName = AppendPath(App.Path) & strFilePart
                        fGraphics(UBound(fGraphics)).bolExtracted = True
                        DoEvents
                    Else
                        strGraphicType = Split(.GraphicExtended, "|")
                        strGraphicLength = Split(.GraphicSize, "|")
                        strGraphicStartPos = Split(.GraphicStartPos, "|")
                        strGraphicTitle = Split(.GraphicTitle, "|")
                        
                        For lngCounter = 0 To UBound(strGraphicType)
                            ReDim Preserve fGraphics(UBound(fGraphics) + 1)
                            fGraphics(UBound(fGraphics)).bytType = ID3V230X.GetGraphicType(strGraphicType(lngCounter))
                            fGraphics(UBound(fGraphics)).dblLength = strGraphicLength(lngCounter)
                            fGraphics(UBound(fGraphics)).dblStartPosition = strGraphicStartPos(lngCounter)
                            fGraphics(UBound(fGraphics)).strTitle = strGraphicTitle(lngCounter)
                            strFileName = .ExtractGraphic(strFile, strGraphicStartPos(lngCounter), strGraphicLength(lngCounter))
                            strFilePart = GetFileName(strFileName)
                            If FileExists(AppendPath(App.Path) & strFilePart) = True Then Kill AppendPath(App.Path) & strFilePart
                            Name strFileName As AppendPath(App.Path) & strFilePart
                            fGraphics(UBound(fGraphics)).strFileName = AppendPath(App.Path) & strFilePart
                            fGraphics(UBound(fGraphics)).bolExtracted = True
                            DoEvents
                        Next
                    End If
                End If
                
                cboPictureType_Click
            End If
        End With
    End If
    
    Exit Sub
    
ErrHan:
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
    cboGenre.Clear
    cboGenre.AddItem ""
    
    For lngCounter = 0 To UBound(strGenres)
        cboGenre.AddItem strGenres(lngCounter)
    Next
End Sub

Private Sub lvwMP3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static lngLastIndex As Long
    
    If lngLastIndex < 0 Then
        lngLastIndex = ColumnHeader.Index - 1
    Else
        If lngLastIndex = ColumnHeader.Index - 1 Then
            lvwMP3.Sorted = False
            If lvwMP3.SortOrder = lvwAscending Then
                lvwMP3.SortOrder = lvwDescending
            Else
                lvwMP3.SortOrder = lvwAscending
            End If
            lvwMP3.Sorted = True
            Exit Sub
        Else
            lngLastIndex = ColumnHeader.Index - 1
        End If
    End If
    
    lvwMP3.Sorted = False
    lvwMP3.SortOrder = lvwAscending
    lvwMP3.SortKey = lngLastIndex
    lvwMP3.Sorted = True
End Sub

Private Sub lvwMP3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lvwMP3_Click
End Sub

Private Sub DisableControls()
    cboDirectories.Enabled = False
    fraFrame.Enabled = False
    lvwMP3.Enabled = False
    fraID3Buttons.Enabled = False
    fraInfo.Enabled = False
    fraID3.Enabled = False
End Sub

Private Sub EnableControls()
    cboDirectories.Enabled = True
    fraFrame.Enabled = True
    lvwMP3.Enabled = True
    fraID3Buttons.Enabled = True
    fraInfo.Enabled = True
    fraID3.Enabled = True
End Sub

Private Sub LoadPictureTypes()
    cboPictureType.Clear
    
    With cboPictureType
        .AddItem "Other"
        .ItemData(.NewIndex) = 0
        .AddItem "32x32 Icon"
        .ItemData(.NewIndex) = 1
        .AddItem "Other Icon"
        .ItemData(.NewIndex) = 2
        .AddItem "Front Cover"
        .ItemData(.NewIndex) = 3
        .AddItem "Back Cover"
        .ItemData(.NewIndex) = 4
        .AddItem "Leaflet Page"
        .ItemData(.NewIndex) = 5
        .AddItem "Media"
        .ItemData(.NewIndex) = 6
        .AddItem "Lead Performer"
        .ItemData(.NewIndex) = 7
        .AddItem "Artist"
        .ItemData(.NewIndex) = 8
        .AddItem "Conductor"
        .ItemData(.NewIndex) = 9
        .AddItem "Band"
        .ItemData(.NewIndex) = 10
        .AddItem "Composer"
        .ItemData(.NewIndex) = 11
        .AddItem "Lyricist"
        .ItemData(.NewIndex) = 12
        .AddItem "Recording Location"
        .ItemData(.NewIndex) = 13
        .AddItem "During Recording"
        .ItemData(.NewIndex) = 14
        .AddItem "During Performance"
        .ItemData(.NewIndex) = 15
        .AddItem "Movie Capture"
        .ItemData(.NewIndex) = 16
        .AddItem "A Bright Colored Fish"
        .ItemData(.NewIndex) = 17
        .AddItem "Illustration"
        .ItemData(.NewIndex) = 18
        .AddItem "Band Logo"
        .ItemData(.NewIndex) = 19
        .AddItem "Publisher Logo"
        .ItemData(.NewIndex) = 20
        .ListIndex = 17
    End With
End Sub

Private Sub ClearControls()
    txtArtist.Text = ""
    txtAlbum.Text = ""
    txtSong.Text = ""
    txtTrack.Text = ""
    txtYear.Text = ""
    cboGenre.Text = ""
    txtComment.Text = ""
    
    txtOriginalArtist.Text = ""
    txtSubtitle.Text = ""
    txtComposer.Text = ""
    txtCopyright.Text = ""
    txtEncodedBy.Text = ""
    txtLanguage.Text = ""
    lstURL.Clear
    txtUnknown.Text = ""
    
    txtLyrics.Text = ""
    txtSyncLyrics.Text = ""
    
    picPicture.Picture = LoadPicture()
    cboPictureType.ListIndex = 17
    txtPictureTitle.Text = ""
End Sub

Private Sub tmrWinAmp_Timer()
    GetWinAMPWindow
    
    If glngWinAmp <> 0 Then
        lblTime.Caption = GetWinAMPPosition
    Else
        lblTime.Caption = "00:00"
    End If
End Sub

Private Sub txtAlbum_GotFocus()
    txtAlbum.SelStart = 0
    txtAlbum.SelLength = Len(txtAlbum.Text)
End Sub

Private Sub txtArtist_GotFocus()
    txtArtist.SelStart = 0
    txtArtist.SelLength = Len(txtArtist.Text)
End Sub

Private Sub txtComment_GotFocus()
    txtComment.SelStart = 0
    txtComment.SelLength = Len(txtComment.Text)
End Sub

Private Sub txtComposer_GotFocus()
    txtComposer.SelStart = 0
    txtComposer.SelLength = Len(txtComposer.Text)
End Sub

Private Sub txtCopyright_GotFocus()
    txtCopyright.SelStart = 0
    txtCopyright.SelLength = Len(txtCopyright.Text)
End Sub

Private Sub txtEncodedBy_GotFocus()
    txtEncodedBy.SelStart = 0
    txtEncodedBy.SelLength = Len(txtEncodedBy.Text)
End Sub

Private Sub txtLanguage_GotFocus()
    txtLanguage.SelStart = 0
    txtLanguage.SelLength = Len(txtLanguage.Text)
End Sub

Private Sub txtOriginalArtist_GotFocus()
    txtOriginalArtist.SelStart = 0
    txtOriginalArtist.SelLength = Len(txtOriginalArtist.Text)
End Sub

Private Sub txtPictureTitle_GotFocus()
    txtPictureTitle.SelStart = 0
    txtPictureTitle.SelLength = Len(txtPictureTitle.Text)
End Sub

Private Sub txtSong_GotFocus()
    txtSong.SelStart = 0
    txtSong.SelLength = Len(txtSong.Text)
End Sub

Private Sub txtSubtitle_GotFocus()
    txtSubtitle.SelStart = 0
    txtSubtitle.SelLength = Len(txtSubtitle.Text)
End Sub

Private Sub txtSyncLyrics_DblClick()
    Dim strInput As String
    
    strInput = InputBox("Offset times by how many seconds?", "")
    If strInput = "" Then Exit Sub
    If IsNumeric(strInput) = False Then Exit Sub
    
    Dim strLines() As String
    Dim lngCounter As Long
    Dim strLine As String
    Dim strTime As String
    Dim strLyrics As String
    
    On Local Error Resume Next
    
    If txtSyncLyrics.Text <> "" Then
        strLines = Split(txtSyncLyrics.Text, vbCrLf)
        For lngCounter = 0 To UBound(strLines)
            strTime = Left(strLines(lngCounter), InStr(strLines(lngCounter), "]"))
            strLine = Right(strLines(lngCounter), Len(strLines(lngCounter)) - Len(strTime))
            strTime = AddTime(strTime, CLng(strInput))
            strLine = strTime & strLine
            If strLyrics = "" Then
                strLyrics = strLine
            Else
                strLyrics = strLyrics & vbCrLf & strLine
            End If
        Next
        txtSyncLyrics.Text = strLyrics
    End If
End Sub

Private Sub txtTrack_GotFocus()
    txtTrack.SelStart = 0
    txtTrack.SelLength = Len(txtTrack.Text)
End Sub

Private Sub txtUnknown_GotFocus()
    txtUnknown.SelStart = 0
    txtUnknown.SelLength = Len(txtUnknown.Text)
End Sub

Private Sub txtYear_GotFocus()
    txtYear.SelStart = 0
    txtYear.SelLength = Len(txtYear.Text)
End Sub

Private Function ReadFile(ByVal strFileName As String) As String
    Dim intFile As Integer
    Dim strLine As String
    Dim strLyrics As String
    
    On Local Error GoTo ErrHan
    
    If FileExists(strFileName) = False Then GoTo ErrHan
    
    intFile = FreeFile
    Open strFileName For Input As intFile
        While Not EOF(intFile)
            Line Input #intFile, strLine
            If strLyrics = "" Then
                strLyrics = strLine
            Else
                strLyrics = strLyrics & vbCrLf & strLine
            End If
        Wend
    Close intFile
    
    ReadFile = strLyrics
    Exit Function
    
ErrHan:
    strLyrics = ""
End Function

Private Function GetFileName(ByVal strFileName As String) As String
    Dim strParts() As String
    
    On Local Error Resume Next
    
    strParts = Split(strFileName, "\")
    GetFileName = strParts(UBound(strParts))
End Function

Private Sub ProcessLine(ByVal strData As String)
    Dim strLines() As String
    Dim lngCounter As Long
    Dim strTime As String
    Dim strLine As String
    Dim strMS() As String
    
    On Local Error Resume Next
    
    If Left(strData, 1) <> "[" Then Exit Sub
    
    strData = Replace(strData, "][", "]|:|[")
    strLines = Split(strData, "|:|")
    
    If UBound(strLines) > 0 Then
        'more than one time slot
        strTime = Left(strLines(UBound(strLines)), InStr(strLines(UBound(strLines)), "]"))
        strMS = Split(strTime, ":")
        strLine = Right(strLines(UBound(strLines)), Len(strLines(UBound(strLines))) - Len(strTime))
        
        ReDim Preserve Lyrics(UBound(Lyrics) + 1)
        Lyrics(UBound(Lyrics)).intMinute = Replace(strMS(0), "[", "")
        Lyrics(UBound(Lyrics)).intSecond = Replace(strMS(1), "]", "")
        Lyrics(UBound(Lyrics)).strLine = strLine
        
        For lngCounter = (UBound(strLines) - 1) To 0 Step -1
            strTime = Left(strLines(lngCounter), InStr(strLines(lngCounter), "]"))
            strMS = Split(strTime, ":")
        
            ReDim Preserve Lyrics(UBound(Lyrics) + 1)
            Lyrics(UBound(Lyrics)).intMinute = Replace(strMS(0), "[", "")
            Lyrics(UBound(Lyrics)).intSecond = Replace(strMS(1), "]", "")
            Lyrics(UBound(Lyrics)).strLine = strLine
        Next
    Else
        'only one time slot
        strTime = Left(strData, InStr(strData, "]"))
        strMS = Split(strTime, ":")
        strLine = Right(strData, Len(strData) - Len(strTime))
        
        ReDim Preserve Lyrics(UBound(Lyrics) + 1)
        Lyrics(UBound(Lyrics)).intMinute = Replace(strMS(0), "[", "")
        Lyrics(UBound(Lyrics)).intSecond = Replace(strMS(1), "]", "")
        Lyrics(UBound(Lyrics)).strLine = strLine
    End If
End Sub

Private Sub DisplayLyrics()
    Dim lngCounter As Long
    
    For lngCounter = 1 To UBound(Lyrics)
        If txtSyncLyrics.Text = "" Then
            txtSyncLyrics.Text = "[" & Format(Lyrics(lngCounter).intMinute, "00") & ":" & Format(Lyrics(lngCounter).intSecond, "00") & "]" & Lyrics(lngCounter).strLine
        Else
            txtSyncLyrics.Text = txtSyncLyrics.Text & vbCrLf & "[" & Format(Lyrics(lngCounter).intMinute, "00") & ":" & Format(Lyrics(lngCounter).intSecond, "00") & "]" & Lyrics(lngCounter).strLine
        End If
    Next
End Sub

Private Sub SortLyrics()
    Dim lngOuterLoop As Long
    Dim lngInnerLoop As Long
    Dim lngTime1 As Long
    Dim lngTime2 As Long
    Dim intFlag As Integer
    
    For lngOuterLoop = 1 To UBound(Lyrics) - 1
        For lngInnerLoop = (lngOuterLoop + 1) To UBound(Lyrics)
            lngTime1 = (Lyrics(lngInnerLoop - 1).intMinute * 60) + (Lyrics(lngInnerLoop - 1).intSecond)
            lngTime2 = (Lyrics(lngInnerLoop).intMinute * 60) + (Lyrics(lngInnerLoop).intSecond)
            If lngTime1 > lngTime2 Then
                MoveItemDown lngInnerLoop - 1
                intFlag = 1
            End If
        Next
        If intFlag = 1 Then
            intFlag = 0
            lngOuterLoop = lngOuterLoop - 1
        End If
    Next
End Sub

Private Sub MoveItemDown(ByVal lngStart As Long)
    Dim Temp As Lyric
    
    Temp.intMinute = Lyrics(lngStart).intMinute
    Temp.intSecond = Lyrics(lngStart).intSecond
    Temp.strLine = Lyrics(lngStart).strLine
    
    Lyrics(lngStart).intMinute = Lyrics(lngStart + 1).intMinute
    Lyrics(lngStart).intSecond = Lyrics(lngStart + 1).intSecond
    Lyrics(lngStart).strLine = Lyrics(lngStart + 1).strLine
    
    Lyrics(lngStart + 1).intMinute = Temp.intMinute
    Lyrics(lngStart + 1).intSecond = Temp.intSecond
    Lyrics(lngStart + 1).strLine = Temp.strLine
End Sub

Private Function AddTime(ByVal strTime As String, ByVal lngAmount As Long)
    Dim strData() As String
    Dim intM As Integer
    Dim intS As Integer
    Dim lngSeconds As Long
    
    strTime = Replace(strTime, "[", "")
    strTime = Replace(strTime, "]", "")
    strData = Split(strTime, ":")
    
    intM = strData(0)
    intS = strData(1)
    
    lngSeconds = (intM * 60) + intS + lngAmount
    AddTime = "[" & MakeTime(lngSeconds) & "]"
End Function

Private Function MakeTime(ByVal lngSeconds As Long) As String
    Dim lngSec As Long
    Dim lngMin As Long
    
    lngMin = lngSeconds \ 60
    lngSec = lngSeconds - (lngMin * 60)
    
    MakeTime = Format(lngMin, "00") & ":" & Format(lngSec, "00")
End Function
