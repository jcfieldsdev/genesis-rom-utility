VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2670
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1842.881
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMegaDriveIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   720
      Left            =   360
      Picture         =   "About.frx":0000
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   5
      Top             =   1080
      Width           =   720
   End
   Begin VB.PictureBox picGenesisIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   720
      Left            =   360
      Picture         =   "About.frx":1F366
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   240
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label lblWebSite 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   3285
   End
   Begin VB.Label lblAuthor 
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   3285
   End
   Begin VB.Label lblTitle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   3285
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const WEB_SITE_URL As String = "https://github.com/jcfieldsdev/genesis-rom-utility"

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub OpenWebSite()
    ' Opens URL in default browser
    ShellExecute 0, "open", WEB_SITE_URL, 0, 0, 1
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    
    lblTitle.Caption = App.Title
    lblAuthor.Caption = "Written by " & App.CompanyName
    lblWebSite.Caption = WEB_SITE_URL
End Sub

Private Sub lblWebSite_Click()
    OpenWebSite
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
