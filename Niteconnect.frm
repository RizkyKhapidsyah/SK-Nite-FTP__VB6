VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNiteConnect 
   BackColor       =   &H8000000A&
   Caption         =   "Nite Connect - FTP "
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   34
      ImageHeight     =   34
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Niteconnect.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Niteconnect.frx":091A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Niteconnect.frx":1294
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Niteconnect.frx":1B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Niteconnect.frx":2394
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   1217
      ButtonWidth     =   1085
      ButtonHeight    =   1058
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connect"
            Description     =   "Connect"
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Disconnect"
            Description     =   "Disconnect"
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Download"
            Description     =   "Download"
            Object.ToolTipText     =   "Download"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Upload"
            Description     =   "Upload"
            Object.ToolTipText     =   "Upload"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Description     =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1920
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RemotePort      =   21
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2640
      Width           =   4245
   End
   Begin VB.Frame fraServerFiles 
      Caption         =   "Server files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   4320
      TabIndex        =   15
      Top             =   1440
      Width           =   4035
      Begin VB.ListBox lisServerFiles 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   4740
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   3765
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3360
      ScaleHeight     =   315
      ScaleWidth      =   765
      TabIndex        =   12
      Top             =   960
      Width           =   825
      Begin VB.CommandButton cmdNil 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "No username and password"
         Top             =   0
         Width           =   225
      End
      Begin VB.CommandButton cmdPrivate 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Use registered userName"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPublic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Use Anonymous as userName"
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame fraLocalFiles 
      Caption         =   "Local files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   4215
      Begin VB.FileListBox File1 
         BackColor       =   &H00E0E0E0&
         Height          =   1845
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3975
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   4005
      Begin VB.Label lblStatus 
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   270
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraLogon 
      BackColor       =   &H8000000A&
      Caption         =   "Log on"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   4245
      Begin VB.TextBox txtport 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txbPassword 
         BackColor       =   &H80000009&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2085
      End
      Begin VB.TextBox txbUserName 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txbURL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   660
         TabIndex        =   0
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Port #:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPassword 
         BackColor       =   &H8000000A&
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblUserName 
         BackColor       =   &H8000000A&
         Caption         =   "UserName:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblURL 
         BackColor       =   &H8000000A&
         Caption         =   "URL:"
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuReconnect 
         Caption         =   "&Reconnect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "&Download"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "&Upload"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAboutabout 
      Caption         =   "&About"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmNiteConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Option Explicit

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags _
      As Long, ByVal dwReserved As Long) As Long


Const defaultURL = "ftp://ftp.microsoft.com"
Const defaultUserName = "fillinyourusername"
Const defaultPassword = "fillinyourpassword"
Const defaultEMailAddress = "fillinyourisp.com"

Dim ConnectedFlag As Boolean
Dim ServerDirFlag As Boolean
Dim DownloadFlag As Boolean
Dim UploadFlag As Boolean
Dim FileSizeFlag As Boolean

Dim homeLen As Integer
Dim LocFilespec As String
Dim SerFilespec As String
Dim gFileSize As String




Private Sub Form_Load()
    GetStartingDefaults
    ConnectedFlag = False
    ClearFlags
 
    txtport.Text = Inet1.RemotePort
    
End Sub



' To access "ftp://ftp.microsoft.com", username and password
' are not required to be typed out.  So default these here.
Private Sub GetStartingDefaults()
    txbURL.Text = defaultURL
    txbUserName.Text = ""
    txbPassword.Text = ""
End Sub



Private Sub cmdConnect_click()
     On Error Resume Next
     Dim tmp As String
     Dim i As Integer
     
     Inet1.Cancel
     Inet1.Execute , "CLOSE"
    
     Err.Clear
     On Error GoTo errHandler
    
     ClearFlags
    
     If Len(txbURL) < 6 Then
          MsgBox "No URL yet"
          Exit Sub
     End If
    
     If UCase(Left(txbURL, 6)) <> "FTP://" Then
          MsgBox "No FTP protocol entered in URL"
          Exit Sub
     End If
    
     lblStatus.Caption = "To connect ...."
    
       ' (Note we use txtURL.Text here; you can just use txtURL if you wish)
     Inet1.AccessType = icUseDefault
     Inet1.URL = LTrim(Trim(txbURL.Text))
     Inet1.UserName = LTrim(Trim(txbUserName.Text))
     Inet1.Password = LTrim(Trim(txbPassword.Text))
     Inet1.RequestTimeout = 40
     
     If txtport.Text <> "" Then
     Inet1.RemotePort = txtport.Text
 End If


            
       ' Will force to bring up Dialup Dialog if not already having a line
     ServerDirFlag = True
     Inet1.Execute , "DIR"
     Do While Inet1.StillExecuting
          DoEvents
          ' Connection not established yet, hence cannot
          ' try to fall back on ConnectedFlag to exit
     Loop
     txbURL.Text = Inet1.URL
     
          ' Home portion
     For i = 7 To Len(txbURL.Text)
          tmp = Mid(txbURL.Text, i, 1)
          If tmp = "/" Then
               Exit For
          End If
     Next i
     homeLen = i - 1
     
     If IsNetConnected() Then
          ConnectedFlag = True
          
     Else
          GoTo errHandler
     End If
     Exit Sub
    
errHandler:
    If icExecuting Then
           ' We place this here in case command for "CLOSE" failed.
           ' With Inet, one can never tell.
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                  lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     ErrMsgProc "cmdConnect_Click"
End Sub



Private Sub cmdExit_Click()
    On Error Resume Next
    Inet1.Execute , "CLOSE"
End
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Inet1.Execute , "CLOSE"
    Unload Me
End Sub



Private Sub cmdPublic_Click()
    txbUserName.Text = "anonymous"
    txbPassword.PasswordChar = ""
    txbPassword.Text = defaultEMailAddress
    txbURL.SetFocus
End Sub



Private Sub cmdPrivate_Click()
    txbUserName.Text = defaultUserName
    txbPassword.PasswordChar = "*"
    txbPassword.Text = defaultPassword
    txbURL.SetFocus
End Sub



Private Sub cmdNil_Click()
    txbUserName.Text = ""
    txbPassword.Text = ""
    txbURL.SetFocus
End Sub




Private Sub cmdDisconnect_Click()
    On Error Resume Next
    Inet1.Cancel
    Inet1.Execute , "CLOSE"
    lblStatus.Caption = "Unconnected"
       ' Put back starting default
    GetStartingDefaults
    ConnectedFlag = False
    lisServerFiles.Clear
    ClearFlags
    
End Sub




Private Sub ClearFlags()
    ServerDirFlag = False
    DownloadFlag = False
    UploadFlag = False
    FileSizeFlag = False
End Sub






Private Sub cmdDownLoad_Click()
     On Error GoTo errHandler
     
     If ConnectedFlag = False Then
          MsgBox "No connection yet"
          Exit Sub
     ElseIf lisServerFiles.ListCount = 0 Then
          MsgBox "No server file listed yet"
          Exit Sub
     ElseIf Right(lisServerFiles.Text, 1) = "/" Then
          MsgBox "Selected item is a directory only." & vbCrLf & vbCrLf & _
             "To list files under that dir, double click on it."
          Exit Sub
     End If
    
     lblStatus.Caption = "Retreiving file..."
     SerFilespec = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & _
               "/" & lisServerFiles.Text
     SerFilespec = Right(SerFilespec, Len(SerFilespec) - 1)
     
        ' Use same file name and store it in current dir of local. Parse
        ' above SerFilespec and take only the file name as LocFileSpec.
     LocFilespec = SerFilespec
     Do While InStr(LocFilespec, "/") <> 0
         LocFilespec = Right(LocFilespec, Len(LocFilespec) - _
              InStr(LocFilespec, "/"))
     Loop
     
     If IsFileThere(LocFilespec) Then
          If MsgBox(LocFilespec & " already exist. Overwrite?", _
               vbYesNo + vbQuestion) = vbNo Then
               Exit Sub
          End If
     End If
     
     lblStatus.Caption = "Requesting for file size..."
     
     gFileSize = ""
     FileSizeFlag = True
     Inet1.Execute , "SIZE " & SerFilespec
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
         
     If gFileSize = "" Then
          MsgBox "Selected file has 0 byte content."
          Exit Sub
     Else
          If MsgBox("File size is " & gFileSize & " bytes." & vbCrLf & vbCrLf & _
                  "Proceed to download?", vbYesNo + vbQuestion) = vbNo Then
              Exit Sub
          End If
     End If
     
     DownloadFlag = True
     Inet1.Execute , "Get " & SerFilespec & " " & LocFilespec
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop

     lblStatus.Caption = "Connected"
     File1.Refresh
     Exit Sub
     
errHandler:
    If icExecuting Then
        If ConnectedFlag = False Then
            Exit Sub
        End If
        
        If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
            Inet1.Cancel
            If Inet1.StillExecuting Then
                lblStatus.Caption = "System failed to cancel job"
            End If
        Else
            Resume
        End If
    End If
    ErrMsgProc "cmdDownLoad_Click"
End Sub




' Assuming you have the appropriate privileges on the server
Private Sub cmdUpLoad_Click()
     On Error GoTo errHandler
     Dim tmpPath As String
     Dim tmpFile As String
     Dim bExist As Boolean
     Dim lFileSize As Long
     Dim i
     
     If ConnectedFlag = False Then
          MsgBox "No connection yet"
          Exit Sub
     ElseIf File1.ListCount = 0 Then
          MsgBox "No local file in current dir yet"
          Exit Sub
     ElseIf Not (Right(lisServerFiles.Text, 1) = "/") Then
          MsgBox "Selected server file item is not a directory"
          Exit Sub
     ElseIf lisServerFiles.Text = "../" Then
          MsgBox "No directory name selected yet"
          Exit Sub
     End If
    
     LocFilespec = tmpPath & File1.List(File1.ListIndex)
     If LocFilespec = "" Then
          MsgBox "No local file selected yet"
          Exit Sub
     End If
     
     lFileSize = FileLen(LocFilespec)
     If MsgBox("File size is " & CStr(lFileSize) & " bytes." & vbCrLf & vbCrLf & _
                  "Proceed to upload?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
     End If
    
     lblStatus.Caption = "Uploading file..."
     
     If Right(Dir1.Path, 1) <> "\" Then
          tmpPath = Dir1.Path & "\"
     Else
          tmpPath = Dir1.Path                   ' e.g. root "C:\"
     End If
     
     SerFilespec = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & _
               "/" & lisServerFiles.Text
          ' Remove the front "/" from above
     SerFilespec = Right(SerFilespec, Len(SerFilespec) - 1)
     
     SerFilespec = SerFilespec & File1.List(File1.ListIndex)
     
          ' In order to test whether same file on server already exists
     lblStatus.Caption = "Verifying existence of file of same name..."
     tmpPath = SerFilespec
     ServerDirFlag = True
     Inet1.Execute , "DIR " & tmpPath & "/*.*"
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
         
     bExist = False
     If lisServerFiles.ListCount > 0 Then
          For i = 0 To lisServerFiles.ListCount - 1
               tmpFile = lisServerFiles.List(i)
               If tmpFile = File1.List(File1.ListIndex) Then
                    bExist = True
                    Exit For
               End If
          Next i
     End If
         
          ' Go back
     ServerDirFlag = True
     Inet1.Execute , "DIR ../*"
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
          
          
     If bExist Then
          If MsgBox("File already exist in selected server dir.  Supersede?", _
                  vbYesNo + vbQuestion) = vbNo Then
               Exit Sub
          End If
     End If
     
     Exit Sub
         
     UploadFlag = True
     Inet1.Execute , "PUT " & LocFilespec & " " & SerFilespec
     
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop

     lblStatus.Caption = "Connected"
     Exit Sub
    
errHandler:
     If icExecuting Then
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                   lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     ErrMsgProc "cmdUpload_Click"
End Sub



Private Sub lbServerFilesHelp_Click()
     MsgBox "Help:" & vbCrLf & vbCrLf & _
          "To change dir, double click a directory item on list." & vbCrLf & _
          "   (To go up one level, click the '../' item)" & vbCrLf & vbCrLf & _
          "To select a file for download, highlight it then" & vbCrLf & _
          "   click Download button (will report file size)." & vbCrLf & vbCrLf & _
          "To upload a local file, highlight a server dir first," & vbCrLf & _
          "   highlight a local file, then click Upload button." & vbCrLf & vbCrLf
End Sub



Private Sub lblLocalFilesHelp_Click()
     MsgBox "Help:" & vbCrLf & vbCrLf & _
          "To see file size of a local file, double click the" & vbCrLf & _
            "   local file item." & vbCrLf & vbCrLf & _
          "For other Help, refer Server Files." & vbCrLf & vbCrLf
End Sub



' For local files, we have FileSystem control to go up and down of dir hierachy
' and list individual files under a dir, but for server files listing, we have to
' provide a similar facility.
Private Sub lisServerFiles_dblClick()
     On Error GoTo errHandler
     
     If Not (Right(lisServerFiles.Text, 1) = "/") Then
          Exit Sub
     End If
     
     Dim tmpDir As String, tmp As String
     Dim i
     If Trim(lisServerFiles.Text) = "../" Then
          For i = Len(txbURL.Text) To 7 Step -1
               tmp = Mid(txbURL.Text, i, 1)
               If tmp = "/" Then
                    Exit For
               End If
          Next i
          If i = 7 Then
               MsgBox "No upper level of dir"
               Exit Sub
          End If
          txbURL.Text = Left(txbURL.Text, i - 1)
             ' Relative dir
          tmpDir = "../*"
     Else
          txbURL.Text = txbURL.Text & "/" & _
                   Left(lisServerFiles.Text, Len(lisServerFiles.Text) - 1)
          tmpDir = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & "/*"
     End If
     ServerDirFlag = True
     Inet1.Execute , "DIR " & tmpDir
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
     Exit Sub
    
errHandler:
    Select Case Err.Number
        Case icExecuting
             Resume
        Case Else
             ErrMsgProc "lisServerFiles_dblClick"
     End Select
End Sub



Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub



Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub



Private Sub File1_dblClick()
    If File1.ListCount = 0 Then
         Exit Sub
    End If
    Dim lFileSize As Long
    lFileSize = FileLen(File1.List(File1.ListIndex))
    MsgBox CStr(lFileSize) & " bytes"
End Sub




Private Sub Inet1_StateChanged(ByVal State As Integer)
    On Error Resume Next
    Select Case State
        Case icError                                      ' 11
            lblStatus = Inet1.ResponseCode & ": " & Inet1.ResponseInfo
            Inet1.Execute , "CLOSE"
            lblStatus.Caption = "Unconnected"
            lisServerFiles.Clear
            ConnectedFlag = False
            ServerDirFlag = False
            DownloadFlag = False
            
            
        Case icResponseCompleted                          ' 12
            Dim bDone As Boolean
            Dim tmpData As Variant       ' GetChunk returns Variant type
            
            If ServerDirFlag = True Then
                 Dim dirData As String
                 Dim strEntry As String
                 Dim i As Integer, k As Integer
            
                 tmpData = Inet1.GetChunk(4096, icString)
                 dirData = dirData & tmpData
            
                 If dirData <> "" Then
                     lisServerFiles.Clear
                       ' Use relative address to allow one dir level up
                     lisServerFiles.AddItem ("../")
                     For i = 1 To Len(dirData) - 1
                          k = InStr(i, dirData, vbCrLf)        ' We don't want CRLF
                          strEntry = Mid(dirData, i, k - i)
                          If Right(strEntry, 1) = "/" Then
                               strEntry = Left(strEntry, Len(strEntry) - 1) & "/"
                          End If
                          If Trim(strEntry) <> "" Then
                               lisServerFiles.AddItem strEntry
                          End If
                          i = k + 1
                          DoEvents
                     Next i
                     lisServerFiles.ListIndex = 0
                 End If
                 
                 ServerDirFlag = False
                 lblStatus.Caption = "Dir completed"
                 
            ElseIf DownloadFlag Then
                 Dim varData As Variant
                 
                 bDone = False

                 Open LocFilespec For Binary Access Write As #1
    
                   ' Get first chunk
                 tmpData = Inet1.GetChunk(10240, icByteArray)
                 DoEvents
                 If Len(tmpData) = 0 Then
                      bDone = True
                 End If
                 Do While Not bDone
                      varData = tmpData
                      Put #1, , varData
                      tmpData = Inet1.GetChunk(10240, icByteArray)
                      DoEvents
                      If ConnectedFlag = False Then
                           Exit Sub
                      End If
                      If Len(tmpData) = 0 Then
                            bDone = True
                      End If
                 Loop
                 Close #1
                 DownloadFlag = False
                 DoEvents
                 lblStatus.Caption = "Download completed"
                 DownloadFlag = False
                 MsgBox "Download completed:" & vbCrLf & vbCrLf & _
                     "File in current dir, named  " & LocFilespec
                 
            ElseIf UploadFlag Then
                 lblStatus.Caption = "Connected"
                 UploadFlag = False
                 MsgBox "Download completed: File in " & LocFilespec
                 
            ElseIf FileSizeFlag Then
                 Dim sizeData As String
            
                 tmpData = Inet1.GetChunk(1024, icString)
                 DoEvents
                 If Len(tmpData) > 0 Then
                      sizeData = sizeData & tmpData
                 End If
                 
                 gFileSize = sizeData
                 FileSizeFlag = False
                 
            Else
                 lblStatus.Caption = "Connected"
            End If
            
            
        Case icNone                                       ' 0
            lblStatus.Caption = "No state to report"
        Case icResolvingHost                              ' 1
            lblStatus.Caption = "Resolving host..."
        Case icHostResolved                               ' 2
            lblStatus.Caption = "Host resolved - found its IP address"
        Case icConnecting                                 ' 3
            lblStatus.Caption = "Connecting..."
        Case icConnected                                  ' 4
            lblStatus.Caption = "Connected"
        Case icRequesting                                 ' 5
            lblStatus.Caption = "Sending requesst..."
        Case icRequestSent                                ' 6
            lblStatus.Caption = "Request sent"
        Case icReceivingResponse                          ' 7
            lblStatus = "Receiving data..."
        Case icResponseReceived                           ' 8
            lblStatus = "Response received"
        Case icDisconnecting                              ' 9
            lblStatus.Caption = "Disconnecting..."
        Case icDisconnected                               '10
            lblStatus = "Disconnected"
    End Select
End Sub



Function IsNetConnected() As Boolean
    IsNetConnected = InternetGetConnectedState(0, 0)
End Function
                  


Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub



Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function

Private Sub mnuAbout_Click()
About.Show
End Sub

Private Sub mnuConnect_Click()
On Error Resume Next
     Dim tmp As String
     Dim i As Integer
     
     Inet1.Cancel
     Inet1.Execute , "CLOSE"
    
     Err.Clear
     On Error GoTo errHandler
    
     ClearFlags
    
     If Len(txbURL) < 6 Then
          MsgBox "No URL yet"
          Exit Sub
     End If
    
     If UCase(Left(txbURL, 6)) <> "FTP://" Then
          MsgBox "No FTP protocol entered in URL"
          Exit Sub
     End If
    
     lblStatus.Caption = "To connect ...."
    
       ' (Note we use txtURL.Text here; you can just use txtURL if you wish)
     Inet1.AccessType = icUseDefault
     Inet1.URL = LTrim(Trim(txbURL.Text))
     Inet1.UserName = LTrim(Trim(txbUserName.Text))
     Inet1.Password = LTrim(Trim(txbPassword.Text))
     Inet1.RequestTimeout = 40
     
     If txtport.Text <> "" Then
     Inet1.RemotePort = txtport.Text
 End If


            
       ' Will force to bring up Dialup Dialog if not already having a line
     ServerDirFlag = True
     Inet1.Execute , "DIR"
     Do While Inet1.StillExecuting
          DoEvents
          ' Connection not established yet, hence cannot
          ' try to fall back on ConnectedFlag to exit
     Loop
     txbURL.Text = Inet1.URL
     
          ' Home portion
     For i = 7 To Len(txbURL.Text)
          tmp = Mid(txbURL.Text, i, 1)
          If tmp = "/" Then
               Exit For
          End If
     Next i
     homeLen = i - 1
     
     If IsNetConnected() Then
          ConnectedFlag = True
          
     Else
          GoTo errHandler
     End If
     Exit Sub
    
errHandler:
    If icExecuting Then
           ' We place this here in case command for "CLOSE" failed.
           ' With Inet, one can never tell.
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                  lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     ErrMsgProc "cmdConnect_Click"

End Sub

Private Sub mnuDisconnect_Click()
On Error Resume Next
    Inet1.Cancel
    Inet1.Execute , "CLOSE"
    lblStatus.Caption = "Unconnected"
       ' Put back starting default
    GetStartingDefaults
    ConnectedFlag = False
    lisServerFiles.Clear
    ClearFlags
    
End Sub

Private Sub mnuDownload_Click()
 On Error GoTo errHandler
     
     If ConnectedFlag = False Then
          MsgBox "No connection yet"
          Exit Sub
     ElseIf lisServerFiles.ListCount = 0 Then
          MsgBox "No server file listed yet"
          Exit Sub
     ElseIf Right(lisServerFiles.Text, 1) = "/" Then
          MsgBox "Selected item is a directory only." & vbCrLf & vbCrLf & _
             "To list files under that dir, double click on it."
          Exit Sub
     End If
    
     lblStatus.Caption = "Retreiving file..."
     SerFilespec = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & _
               "/" & lisServerFiles.Text
     SerFilespec = Right(SerFilespec, Len(SerFilespec) - 1)
     
        ' Use same file name and store it in current dir of local. Parse
        ' above SerFilespec and take only the file name as LocFileSpec.
     LocFilespec = SerFilespec
     Do While InStr(LocFilespec, "/") <> 0
         LocFilespec = Right(LocFilespec, Len(LocFilespec) - _
              InStr(LocFilespec, "/"))
     Loop
     
     If IsFileThere(LocFilespec) Then
          If MsgBox(LocFilespec & " already exist. Overwrite?", _
               vbYesNo + vbQuestion) = vbNo Then
               Exit Sub
          End If
     End If
     
     lblStatus.Caption = "Requesting for file size..."
     
     gFileSize = ""
     FileSizeFlag = True
     Inet1.Execute , "SIZE " & SerFilespec
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
         
     If gFileSize = "" Then
          MsgBox "Selected file has 0 byte content."
          Exit Sub
     Else
          If MsgBox("File size is " & gFileSize & " bytes." & vbCrLf & vbCrLf & _
                  "Proceed to download?", vbYesNo + vbQuestion) = vbNo Then
              Exit Sub
          End If
     End If
     
     DownloadFlag = True
     Inet1.Execute , "Get " & SerFilespec & " " & LocFilespec
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop

     lblStatus.Caption = "Connected"
     File1.Refresh
     Exit Sub
     
errHandler:
    If icExecuting Then
        If ConnectedFlag = False Then
            Exit Sub
        End If
        
        If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
            Inet1.Cancel
            If Inet1.StillExecuting Then
                lblStatus.Caption = "System failed to cancel job"
            End If
        Else
            Resume
        End If
    End If
    ErrMsgProc "cmdDownLoad_Click"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuReconnect_Click()
On Error Resume Next
     Dim tmp As String
     Dim i As Integer
     
     Inet1.Cancel
     Inet1.Execute , "CLOSE"
    
     Err.Clear
     On Error GoTo errHandler
    
     ClearFlags
    
     If Len(txbURL) < 6 Then
          MsgBox "No URL yet"
          Exit Sub
     End If
    
     If UCase(Left(txbURL, 6)) <> "FTP://" Then
          MsgBox "No FTP protocol entered in URL"
          Exit Sub
     End If
    
     lblStatus.Caption = "To connect ...."
    
       ' (Note we use txtURL.Text here; you can just use txtURL if you wish)
     Inet1.AccessType = icUseDefault
     Inet1.URL = LTrim(Trim(txbURL.Text))
     Inet1.UserName = LTrim(Trim(txbUserName.Text))
     Inet1.Password = LTrim(Trim(txbPassword.Text))
     Inet1.RequestTimeout = 40
     
     If txtport.Text <> "" Then
     Inet1.RemotePort = txtport.Text
 End If


            
       ' Will force to bring up Dialup Dialog if not already having a line
     ServerDirFlag = True
     Inet1.Execute , "DIR"
     Do While Inet1.StillExecuting
          DoEvents
          ' Connection not established yet, hence cannot
          ' try to fall back on ConnectedFlag to exit
     Loop
     txbURL.Text = Inet1.URL
     
          ' Home portion
     For i = 7 To Len(txbURL.Text)
          tmp = Mid(txbURL.Text, i, 1)
          If tmp = "/" Then
               Exit For
          End If
     Next i
     homeLen = i - 1
     
     If IsNetConnected() Then
          ConnectedFlag = True
          
     Else
          GoTo errHandler
     End If
     Exit Sub
    
errHandler:
    If icExecuting Then
           ' We place this here in case command for "CLOSE" failed.
           ' With Inet, one can never tell.
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                  lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     ErrMsgProc "cmdConnect_Click"

End Sub

Private Sub mnuUpload_Click()
  On Error GoTo errHandler
     Dim tmpPath As String
     Dim tmpFile As String
     Dim bExist As Boolean
     Dim lFileSize As Long
     Dim i
     
     If ConnectedFlag = False Then
          MsgBox "No connection yet"
          Exit Sub
     ElseIf File1.ListCount = 0 Then
          MsgBox "No local file in current dir yet"
          Exit Sub
     ElseIf Not (Right(lisServerFiles.Text, 1) = "/") Then
          MsgBox "Selected server file item is not a directory"
          Exit Sub
     ElseIf lisServerFiles.Text = "../" Then
          MsgBox "No directory name selected yet"
          Exit Sub
     End If
    
     LocFilespec = tmpPath & File1.List(File1.ListIndex)
     If LocFilespec = "" Then
          MsgBox "No local file selected yet"
          Exit Sub
     End If
     
     lFileSize = FileLen(LocFilespec)
     If MsgBox("File size is " & CStr(lFileSize) & " bytes." & vbCrLf & vbCrLf & _
                  "Proceed to upload?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
     End If
    
     lblStatus.Caption = "Uploading file..."
     
     If Right(Dir1.Path, 1) <> "\" Then
          tmpPath = Dir1.Path & "\"
     Else
          tmpPath = Dir1.Path                   ' e.g. root "C:\"
     End If
     
     SerFilespec = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & _
               "/" & lisServerFiles.Text
          ' Remove the front "/" from above
     SerFilespec = Right(SerFilespec, Len(SerFilespec) - 1)
     
     SerFilespec = SerFilespec & File1.List(File1.ListIndex)
     
          ' In order to test whether same file on server already exists
     lblStatus.Caption = "Verifying existence of file of same name..."
     tmpPath = SerFilespec
     ServerDirFlag = True
     Inet1.Execute , "DIR " & tmpPath & "/*.*"
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
         
     bExist = False
     If lisServerFiles.ListCount > 0 Then
          For i = 0 To lisServerFiles.ListCount - 1
               tmpFile = lisServerFiles.List(i)
               If tmpFile = File1.List(File1.ListIndex) Then
                    bExist = True
                    Exit For
               End If
          Next i
     End If
         
          ' Go back
     ServerDirFlag = True
     Inet1.Execute , "DIR ../*"
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
          
          
     If bExist Then
          If MsgBox("File already exist in selected server dir.  Supersede?", _
                  vbYesNo + vbQuestion) = vbNo Then
               Exit Sub
          End If
     End If
     
     Exit Sub
         
     UploadFlag = True
     Inet1.Execute , "PUT " & LocFilespec & " " & SerFilespec
     
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop

     lblStatus.Caption = "Connected"
     Exit Sub
    
errHandler:
     If icExecuting Then
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                   lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     ErrMsgProc "cmdUpload_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
    Case "Connect"
        cmdConnect_click
    Case "Disconnect"
        cmdDisconnect_Click
    Case "Download"
        cmdDownLoad_Click
    Case "upload"
        cmdUpLoad_Click
    Case "Exit"
        cmdExit_Click
        End Select


End Sub
