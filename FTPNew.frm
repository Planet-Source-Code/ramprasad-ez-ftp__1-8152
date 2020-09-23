VERSION 5.00
Object = "{6580F760-7819-11CF-B86C-444553540000}#1.0#0"; "EZFTP.OCX"
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EZ FTP"
   ClientHeight    =   5355
   ClientLeft      =   2250
   ClientTop       =   1650
   ClientWidth     =   6990
   HelpContextID   =   560
   Icon            =   "FTPNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5355
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   3270
      Top             =   3210
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   510
      ScaleHeight     =   195
      ScaleWidth      =   5835
      TabIndex        =   24
      Top             =   4560
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton AbortTransfer 
      Caption         =   "A&bort"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4650
      TabIndex        =   21
      Top             =   4890
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Exit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5880
      TabIndex        =   20
      Top             =   4890
      Width           =   1035
   End
   Begin VB.OptionButton ASCIIMode 
      Caption         =   "ASCII"
      Height          =   195
      Left            =   3540
      TabIndex        =   19
      Top             =   4980
      Width           =   855
   End
   Begin VB.OptionButton BinaryMode 
      Caption         =   "Binary"
      Height          =   195
      Left            =   2520
      TabIndex        =   18
      Top             =   4980
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton ToRemote 
      Caption         =   "-->"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1980
      Width           =   495
   End
   Begin VB.CommandButton ToLocal 
      Caption         =   "<--"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.Frame Remote 
      Caption         =   "Remote System"
      Height          =   4335
      Left            =   3840
      TabIndex        =   1
      Top             =   60
      Width           =   3075
      Begin VB.CommandButton cmdremoteview 
         Caption         =   "Exec"
         Height          =   315
         Left            =   2460
         TabIndex        =   23
         Top             =   2220
         Width           =   495
      End
      Begin VB.CommandButton RemoteDEL 
         Caption         =   "DEL"
         Height          =   315
         Left            =   2460
         TabIndex        =   17
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton RemoteMD 
         Caption         =   "MD"
         Height          =   315
         Left            =   2460
         TabIndex        =   15
         Top             =   1380
         Width           =   495
      End
      Begin VB.CommandButton RemoteRD 
         Caption         =   "RD"
         Height          =   315
         Left            =   2460
         TabIndex        =   14
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton RemoteCD 
         Caption         =   "CD"
         Height          =   315
         Left            =   2460
         TabIndex        =   13
         Top             =   660
         Width           =   495
      End
      Begin VB.ListBox RemoteFiles 
         Height          =   2400
         Left            =   180
         TabIndex        =   12
         Top             =   1800
         Width           =   2235
      End
      Begin VB.ListBox RemoteDirectories 
         Height          =   1035
         Left            =   180
         TabIndex        =   11
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label RemotePWD 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   2715
      End
   End
   Begin VB.Frame Local 
      Caption         =   "Local System"
      Height          =   4335
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3075
      Begin VB.CommandButton cmdlocalview 
         Caption         =   "Exec"
         Height          =   315
         Left            =   2460
         TabIndex        =   22
         Top             =   2220
         Width           =   495
      End
      Begin VB.CommandButton LocalDEL 
         Caption         =   "DEL"
         Height          =   315
         Left            =   2460
         TabIndex        =   16
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton LocalMD 
         Caption         =   "MD"
         Height          =   315
         Left            =   2460
         TabIndex        =   10
         Top             =   1380
         Width           =   495
      End
      Begin VB.CommandButton LocalRD 
         Caption         =   "RD"
         Height          =   315
         Left            =   2460
         TabIndex        =   9
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton LocalCD 
         Caption         =   "CD"
         Height          =   315
         Left            =   2460
         TabIndex        =   8
         Top             =   660
         Width           =   495
      End
      Begin VB.ListBox LocalFiles 
         Height          =   2400
         Left            =   180
         TabIndex        =   7
         Top             =   1800
         Width           =   2235
      End
      Begin VB.ListBox LocalDirectories 
         Height          =   1035
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label LocalPWD 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   2715
      End
   End
   Begin EZFTPLib.EZFTP FTP 
      Left            =   3240
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   800
      _ExtentY        =   800
      _StockProps     =   0
      LocalFile       =   ""
      RemoteFile      =   ""
      RemoteAddres    =   ""
      UserName        =   ""
      Password        =   ""
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LocalDir As String
Dim i As Integer
Dim remoteview As String
Dim remoteexec As String
'Dim getflag As String
Dim tolocalclick As String
Dim toremoteclick As String
Dim remotedelclick As String
Dim localdelclick As String
Dim AbortedFlag As Boolean
Dim FirstTime As Boolean
Dim NewDirectory As String

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal str As String, ByVal len1 As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Sub RefreshAll()

    RefreshLocal
    RefreshRemote
    
End Sub


Sub RefreshLocal()

    Screen.MousePointer = vbHourglass

'Local Directories and Files
Dim NextLocal As String
Dim FullSpec As String
    
    LocalPWD.Caption = CurDir()

    LocalDirectories.Clear
    LocalFiles.Clear
'    LocalDirectories.Sorted = False
'    LocalFiles.Sorted = False
    If Len(CurDir()) = 3 Then
        FullSpec = CurDir() & "*.*"
    Else
        FullSpec = CurDir() & "\*.*"
    End If
    NextLocal = Dir(FullSpec, vbDirectory + vbNormal)
    Do While NextLocal <> ""
        If Len(CurDir()) = 3 Then
            FullSpec = CurDir() & NextLocal
        Else
            FullSpec = CurDir() & "\" & NextLocal
        End If
        On Error Resume Next
        If (GetAttr(FullSpec) And vbDirectory) = vbDirectory Then
            LocalDirectories.AddItem NextLocal
        Else
            LocalFiles.AddItem NextLocal
        End If
        NextLocal = Dir
    Loop
'    LocalDirectories.Sorted = True
'    LocalFiles.Sorted = True
    
    Screen.MousePointer = vbDefault

End Sub

Sub RefreshRemote()
    
    Screen.MousePointer = vbHourglass
    
'Remote Directories and Files (done in the FTP NextDirectoryEntry event)
    RemotePWD.Caption = FTP.RemoteDirectory
    
    RemoteDirectories.Clear
    RemoteFiles.Clear
    On Error Resume Next
    FTP.GetDirectory ("*.*")

    Screen.MousePointer = vbDefault

End Sub

Private Sub AbortTransfer_Click()

    FTP.AbortTransfer = True
    AbortedFlag = True
    
End Sub


Private Sub ASCIIMode_Click()

    If ASCIIMode.Value = True Then
        FTP.Binary = False
    End If
    
End Sub

Private Sub BinaryMode_Click()

    If BinaryMode.Value = True Then
        FTP.Binary = True
    End If
    
End Sub

Private Sub cmdlocalview_Click()
    Call ShellExecute(hwnd, "Open", LocalPWD & "\" & LocalFiles.Text, "", App.Path, 1)
End Sub

Private Sub cmdremoteview_Click()
    remoteview = "yes"
     ToLocal.Value = 1
    Call ShellExecute(hwnd, "Open", LocalPWD & "\" & RemoteFiles.Text, "", App.Path, 1)
    ChDir LocalDir
    RefreshLocal
    remoteview = ""

End Sub

Private Sub Exit_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()

    If FirstTime = False Then
        FirstTime = True
        If FTP.ProfessionalEdition = True Then
            AbortTransfer.Visible = True
            'ProgressBar.Visible = True
        End If
        DoEvents ' give the form a chance to paint
        RefreshAll
    End If
        
End Sub

Private Sub Form_Load()
    
    Me.Move (Screen.Width \ 2) - (Me.Width \ 2), (Screen.Height \ 2) - (Me.Height \ 2)
    FirstTime = False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    FTP.Disconnect
    
End Sub



Private Sub FTP_NextDirectoryEntry(ByVal FileName As String, ByVal Attributes As Long, ByVal Length As Double)

    If (Attributes And 16) = 16 Or Attributes = 0 Then
        RemoteDirectories.AddItem FileName
    Else
        RemoteFiles.AddItem FileName
    End If
    
End Sub

Private Sub FTP_TransferProgress(ByVal BytesTransferred As Long, ByVal TotalBytes As Long)

'Only fires in the professional editon
'    If ProgressBar.Max <> TotalBytes Then
'        ProgressBar.Max = TotalBytes
'    End If
'    ProgressBar.Value = BytesTransferred
    DoEvents ' give abort a chance...

End Sub

Private Sub LocalCD_Click()

Dim NewDirectory As String

    NewDirectory = InputBox$("Enter directory to change to")
    If NewDirectory = "" Then
        Exit Sub
    End If

    On Error Resume Next
    ChDir NewDirectory
    If Err <> 0 Then
        MsgBox "Unable to change directory", vbExclamation
    Else
        RefreshLocal
    End If
    
End Sub

Private Sub LocalDEL_Click()
    
    If LocalFiles.ListIndex = -1 Then
        Beep
        Exit Sub
    End If
    
    On Error Resume Next
    Kill LocalFiles.Text
    If Err <> 0 Then
        MsgBox "Unable to delete local file", vbExclamation
    Else
        RefreshLocal
    End If

End Sub

Private Sub LocalDirectories_DblClick()

    If LocalDirectories.ListIndex = -1 Then
        Beep
        Exit Sub
    End If
    ChDir LocalDirectories.Text
    RefreshLocal
    
End Sub


Private Sub LocalFiles_DblClick()

    ToRemote.Value = 1
    
End Sub


Private Sub LocalMD_Click()

Dim NewDirectory As String
    
    NewDirectory = InputBox$("Enter new directory name")
    If NewDirectory = "" Then
        Exit Sub
    End If
    
    On Error Resume Next
    MkDir NewDirectory
    If Err <> 0 Then
        MsgBox "Unable to make local directory", vbExclamation
    Else
        RefreshLocal
    End If

End Sub

Private Sub LocalRD_Click()
    
    If LocalDirectories.ListIndex = -1 Then
        Beep
        Exit Sub
    End If
    
    On Error Resume Next
    RmDir LocalDirectories.Text
    If Err <> 0 Then
        MsgBox "Unable to remove local directory", vbExclamation
    Else
        RefreshLocal
    End If

End Sub
Private Sub RemoteCD_Click()

Dim NewDirectory As String

    NewDirectory = InputBox$("Enter directory to change to")
    If NewDirectory = "" Then
        Exit Sub
    End If
    
    On Error Resume Next
    FTP.RemoteDirectory = NewDirectory
    If Err <> 0 Then
        MsgBox "Unable to change directory", vbExclamation
    Else
        RefreshRemote
    End If
        
End Sub

Private Sub RemoteDEL_Click()

    If RemoteFiles.ListIndex = -1 Then
        Beep
        Exit Sub
    End If
    
    On Error Resume Next
    FTP.DeleteFile RemoteFiles.Text
    If Err <> 0 Then
        MsgBox "Unable to delete remote file", vbExclamation
    Else
        RefreshRemote
    End If
    
End Sub

Private Sub RemoteDirectories_DblClick()

    On Error Resume Next
    FTP.RemoteDirectory = RemoteDirectories.Text
    If Err <> 0 Then
        MsgBox "Unable to change directory", vbExclamation
    Else
        RefreshRemote
    End If
    
End Sub


Private Sub RemoteFiles_DblClick()

    ToLocal.Value = 1
    
End Sub


Private Sub RemoteMD_Click()

Dim NewDirectory As String
    
    NewDirectory = InputBox$("Enter new directory name")
    If NewDirectory = "" Then
        Exit Sub
    End If
    
    On Error Resume Next
    FTP.MkDir NewDirectory
    If Err <> 0 Then
        MsgBox "Unable to make remote directory", vbExclamation
    Else
        RefreshRemote
    End If
    

End Sub

Private Sub RemoteRD_Click()

    If RemoteDirectories.ListIndex = -1 Then
        Beep
        Exit Sub
    End If
    
    On Error Resume Next
    FTP.RmDir RemoteDirectories.Text
    If Err <> 0 Then
        MsgBox "Unable to remove remote directory", vbExclamation
    Else
        RefreshRemote
    End If
    
End Sub

Private Sub Timer1_Timer()
    If tolocalclick = "yes" Or toremoteclick = "yes" Or localdelclick = "yes" Or remotedelclick = "yes" Then
        
        Picture1.Visible = True
        Picture1.ForeColor = RGB(0, 0, 255) 'use blue bar
        For i = 0 To 100 Step 2
        updateprogress Picture1, i
        Next
        Picture1.Cls 'clear bar at they end
        Picture1.Visible = False
    
'        Dim Counter As Integer
'        Dim Workarea(250) As String
'        ProgressBar.Min = LBound(Workarea)
'        ProgressBar.Max = UBound(Workarea)
'        ProgressBar.Visible = True
'        'Set the Progress's Value to Min.
'        ProgressBar.Value = ProgressBar.Min
'        'Loop through the array.
'
'        For Counter = LBound(Workarea) To UBound(Workarea)
'            'Set initial values for each item in the array.
'            Workarea(Counter) = "Initial value" & Counter
'            ProgressBar.Value = Counter
'        Next Counter
'        ProgressBar.Visible = False
'
'        ProgressBar.Value = ProgressBar.Min
    End If
End Sub

Private Sub ToLocal_Click()
    
    If remoteview <> "yes" Then
        Dim filestr As String
        filestr = LocalPWD + "\" + RemoteFiles.Text
        If FileExists(filestr) = False Then
            GoTo 0
        Else
            If MsgBox("Do you wish to replace it?", vbYesNo, "File already exists") = vbYes Then
                GoTo 0
            Else
                Exit Sub
            End If
        End If
    End If
0:
    tolocalclick = "yes"
    Timer1_Timer
    If RemoteFiles.ListIndex = -1 Then
        Beep
        Exit Sub
    End If
    If remoteview = "yes" Or remoteexec = "yes" Then
        LocalDir = LocalPWD
        On Error Resume Next
        MkDir LocalPWD & "\temp"
        ChDir LocalPWD & "\temp"
        RefreshLocal
    End If
    FTP.RemoteFile = RemoteFiles.Text
    FTP.LocalFile = RemoteFiles.Text
    Screen.MousePointer = vbHourglass
    AbortTransfer.Enabled = True
    'ProgressBar.Value = 0
    On Error Resume Next
    FTP.GetFile
    AbortTransfer.Enabled = False
    'ProgressBar.Value = 0
    Screen.MousePointer = vbDefault
    If Err <> 0 Then
        MsgBox "Unable to transfer from remote system", vbExclamation
    Else
        If AbortedFlag = True Then
            AbortedFlag = False
            Kill RemoteFiles.Text
        End If
        Beep
        RefreshLocal
    End If
End Sub


Private Sub ToRemote_Click()
    i = 0
    While i < LocalFiles.ListCount
        If RemoteFiles.Text = LocalFiles.List(i) Then
            If MsgBox("Do u wish to replace...", vbQuestion + vbYesNo, "File Already Exists") = vbYes Then
                GoTo 0
            End If
            Exit Sub
        End If
        i = i + 1
    Wend
0:
    toremoteclick = "yes"
    If LocalFiles.ListIndex = -1 Then
        Beep
        Exit Sub
    End If
 
    FTP.LocalFile = LocalFiles.Text
    FTP.RemoteFile = LocalFiles.Text
    Screen.MousePointer = vbHourglass
    AbortTransfer.Enabled = True
'    ProgressBar.Value = 0
    On Error Resume Next
    FTP.PutFile
    AbortTransfer.Enabled = False
    'ProgressBar.Value = 0
    Screen.MousePointer = vbDefault
    If Err <> 0 Then
        MsgBox "Unable to transfer to remote system", vbExclamation
    Else
        If AbortedFlag = True Then
            AbortedFlag = False
            FTP.DeleteFile LocalFiles.Text
        End If
        Beep
        RefreshRemote
    End If
    
End Sub


Public Function FileExists(strfile) As Boolean
    'Checks for the existence of a file by attempting an OPEN.
    Dim intFileNumber
    'get next available file number
    intFileNumber = FreeFile
    On Error Resume Next
    Open strfile For Input As intFileNumber
        If Err = 0 Then
            FileExists = True
        Else
            FileExists = False
        End If
    Close intFileNumber
End Function

Sub updateprogress(pb As Control, ByVal percent)
If tolocalclick = "yes" Or toremoteclick = "yes" Or localdelclick = "yes" Or remotedelclick = "yes" Then
    Dim num$ 'use percent
    If Not pb.AutoRedraw Then 'picture in memory ?
    pb.AutoRedraw = -1 'no, make one
    End If
    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new sclaemodus
    pb.DrawMode = 10 'not XOR Pen Modus
    num$ = Format$(percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    pb.Print num$ 'print percent
    pb.Line (0, 0)-(percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
End If
End Sub


