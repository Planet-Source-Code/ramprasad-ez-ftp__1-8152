VERSION 5.00
Begin VB.Form FTPLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FTP Login"
   ClientHeight    =   2625
   ClientLeft      =   2970
   ClientTop       =   3225
   ClientWidth     =   5925
   HelpContextID   =   10
   Icon            =   "FTPLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Anonymous 
      Alignment       =   1  'Right Justify
      Caption         =   "Anonymous"
      Height          =   255
      Left            =   540
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4740
      TabIndex        =   8
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   4740
      TabIndex        =   7
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1620
      Width           =   2955
   End
   Begin VB.TextBox UserName 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1260
      Width           =   2955
   End
   Begin VB.TextBox Host 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   900
      Width           =   2955
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   540
      TabIndex        =   2
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Host:"
      Height          =   195
      Left            =   540
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "EZ FTP "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FTPLogin.frx":0442
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FTPLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Anonymous_Click()
    If Anonymous.Value = False Then
        UserName.Text = ""
        Password.Text = ""
    Else
        UserName.Text = "anonymous"
        Password.Text = "user@domain.com"
    End If
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()
'*** Code added by HelpWriter ***
    'SetApphelp Me.hwnd
'***********************************

    Me.Move (Screen.Width \ 2) - (Me.Width \ 2), (Screen.Height \ 2) - (Me.Height \ 2)
    
End Sub

Private Sub OK_Click()

    OK.Enabled = False
    Cancel.Enabled = False
    Load Main
    Main.FTP.RemoteAddress = Host.Text
    Main.FTP.UserName = UserName.Text
    Main.FTP.Password = Password.Text
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Main.FTP.Connect
    Screen.MousePointer = vbDefault
    If Err <> 0 Then
        MsgBox "Unable to connect to the specified host", vbCritical
        Unload Main
        Unload Me
    Else
        Main.Show
        Unload Me
    End If
    
End Sub


Sub Form_Unload(Cancel As Integer)
'*** Code added by HelpWriter ***
'*** Subroutine added by HelpWriter ***
    'QuitHelp
'***********************************
End Sub
