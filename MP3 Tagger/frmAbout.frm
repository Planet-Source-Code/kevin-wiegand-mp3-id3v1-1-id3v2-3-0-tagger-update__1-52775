VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MP3 Tagger"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblWebsite 
      AutoSize        =   -1  'True
      Caption         =   "Webiste:  http://a.domaindlx.com/GreatPumpkinator"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Tag             =   "http://a.domaindlx.com/GreatPumpkinator"
      Top             =   1440
      Width           =   3720
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail:  KevinAllenWiegand@hotmail.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Tag             =   "mailto:KevinAllenWiegand@hotmail.com"
      Top             =   1200
      Width           =   2970
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Programmed By:  Kevin Wiegand"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2340
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Version 2.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MP3 Tagger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1725
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SHOWNORMAL = 1

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub RunProgram(ByVal strCommandLineSetting As String, Optional ByVal strParameterString As String)
    ShellExecute 0, "Open", strCommandLineSetting, strParameterString, vbNullString, SHOWNORMAL
End Sub

Private Sub lblEMail_Click()
    RunProgram lblEMail.Tag
End Sub

Private Sub lblWebsite_Click()
    RunProgram lblWebsite.Tag
End Sub
