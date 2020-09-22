VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XCyteMail Network Server"
   ClientHeight    =   5280
   ClientLeft      =   1965
   ClientTop       =   3450
   ClientWidth     =   8760
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8760
   Begin VB.CommandButton Command8 
      Caption         =   "Debug Refresh"
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   -510
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer ShutdownTimer 
      Enabled         =   0   'False
      Left            =   3360
      Top             =   360
   End
   Begin VB.Timer expTimer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4200
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2640
      Top             =   2880
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   8865
      Pattern         =   "*.txt"
      TabIndex        =   17
      Top             =   4635
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2820
      Top             =   3585
   End
   Begin VB.FileListBox File2 
      Height          =   480
      Left            =   8805
      Pattern         =   "*.txt"
      TabIndex        =   8
      Top             =   4065
      Visible         =   0   'False
      Width           =   1350
   End
   Begin XCyteMail.MX MX 
      Left            =   1755
      Top             =   4890
      _ExtentX        =   847
      _ExtentY        =   582
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Server diagnostics"
      Height          =   495
      Left            =   6765
      TabIndex        =   7
      Top             =   7410
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "User manager"
      Height          =   495
      Left            =   6765
      TabIndex        =   6
      Top             =   6930
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Administration"
      Height          =   495
      Left            =   6765
      TabIndex        =   5
      Top             =   6450
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pause"
      Height          =   255
      Left            =   7905
      TabIndex        =   4
      Top             =   4935
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Flush Accounts"
      Height          =   495
      Left            =   6765
      TabIndex        =   3
      Top             =   5970
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Shutdown"
      Height          =   495
      Left            =   6765
      TabIndex        =   2
      Top             =   5490
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Online/Offline"
      Height          =   495
      Left            =   6765
      TabIndex        =   1
      Top             =   5355
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock http 
      Index           =   0
      Left            =   1905
      Top             =   2445
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3040
   End
   Begin MSWinsockLib.Winsock pop3 
      Left            =   1935
      Top             =   1245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   110
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   1920
      Top             =   1875
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   25
   End
   Begin VB.TextBox SvrLog 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   5055
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6885
   End
   Begin VB.PictureBox SideBar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6510
      Left            =   0
      ScaleHeight     =   6510
      ScaleWidth      =   1830
      TabIndex        =   9
      Top             =   -315
      Width           =   1830
      Begin VB.PictureBox Popper 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   1305
         Picture         =   "frmMain.frx":1A7A
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   10
         Tag             =   "0"
         Top             =   2625
         Width           =   450
      End
      Begin VB.Label SideObject 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnostics"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   5
         Left            =   45
         MouseIcon       =   "frmMain.frx":2584
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2385
         Width           =   1485
      End
      Begin VB.Label SideObject 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inboxes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   4
         Left            =   45
         MouseIcon       =   "frmMain.frx":288E
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1980
         Width           =   1485
      End
      Begin VB.Label SideObject 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   3
         Left            =   45
         MouseIcon       =   "frmMain.frx":2B98
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1590
         Width           =   1485
      End
      Begin VB.Label SideObject 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flush"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   2
         Left            =   45
         MouseIcon       =   "frmMain.frx":2EA2
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1185
         Width           =   1485
      End
      Begin VB.Label SideObject 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   45
         MouseIcon       =   "frmMain.frx":31AC
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   795
         Width           =   1485
      End
      Begin VB.Label SideObject 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Online/Offline"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   45
         MouseIcon       =   "frmMain.frx":34B6
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   420
         Width           =   1485
      End
      Begin VB.Image Image1 
         Height          =   6345
         Left            =   0
         Picture         =   "frmMain.frx":37C0
         Stretch         =   -1  'True
         Top             =   -390
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public obj As New Dictionary

Private sentemails As New Collection
Private pop3connections As New Collection


Private Sub Command2_Click()
ShutdownServer
End Sub

Private Sub Command3_Click()
If MsgBox("Are you sure you want to flush all user email accounts?", vbCritical + vbOKCancel, "Confirmation") = vbOK Then
   'FlushAccounts 'Flush to trash backup
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Pause" Then
   ActiveLogging = False
   Command4.Caption = "Log"
Else
   ActiveLogging = True
   Command4.Caption = "Pause"
End If
End Sub

Private Sub Command5_Click()
frmAdmin.Show
End Sub

Private Sub Command6_Click()
frmUsers.Show
frmUsers.RefreshUsers
End Sub

Private Sub Command7_Click()
frmMain.Show
End Sub

Private Sub Command8_Click()
frmMain.Timer1.Enabled = True
End Sub

Private Sub expTimer1_Timer()
If GetHandle <> SideBar.hWnd Then
    Popin
End If
expTimer1.Enabled = False
End Sub

Private Sub File1_PathChange()
frmDiag.File1.Path = File1.Path
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
   If ControlLogging = False Then
   StartService "controllog"
   Else
   StopService "controllog"
   End If
   AppendLog ">"
End If
If KeyCode = vbKeyF6 Then
   If ActiveLogging = False Then
   StartService "activelog"
   Else
   StopService "activelog"
   End If
   AppendLog ">"
End If

If KeyCode = vbKeyF7 Then
   StopShutdown
End If
If KeyCode = vbKeyF8 Then
   If ErrorLog.Tag <> 1 Then
      ErrorLog.Show
      ErrorLog.Left = 0
      ErrorLog.Top = 0
      frmMain.SetFocus
      ErrorLog.Tag = 1
   Else
      ErrorLog.Tag = 0
      ErrorLog.Hide
   End If
End If
If KeyCode = vbKeyF9 Then
   If MsgBox("This feature will (by force), terminate the server application." & vbCrLf & "Are you sure?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
      End
   End If
End If
If KeyCode = vbKeyF2 Then
   If POP = False Then
   StartService "POP3"
   Else
   StopService "POP3"
   End If
   AppendLog ">"
End If
If KeyCode = vbKeyF3 Then
   If SMTP = False Then
   StartService "SMTP"
   Else
   StopService "SMTP"
   End If
   AppendLog ">"
End If
If KeyCode = vbKeyF4 Then
   If WEBMAIL = False Then
   StartService "WEBMAIL"
   Else
   StopService "WEBMAIL"
   End If
   AppendLog ">"
End If
If KeyCode = vbKeyF1 Then
   Dim OriginalLog1 As Boolean
   OriginalLog1 = ControlLogging
   AppendLog "Restarting all services..."
   StartSilent = True
   StopService "POP3"
   StartSilent = True
   StopService "SMTP"
   StartSilent = True
   StopService "WEBMAIL"
   StartSilent = True
   StopService "Activelog"
   StartSilent = True
   StopService "Controllog"
   StartSilent = True
   StartService "POP3"
   StartSilent = True
   StartService "SMTP"
   StartSilent = True
   StartService "WEBMAIL"
   StartSilent = True
   StartService "Activelog"
   StartSilent = True
   StartService "Controllog"
   If OriginalLog1 = False Then ControlLogging = False
   AppendOK
'   AppendLog ">"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case x
Case 7725
    frmMain.Show
    Exit Sub
Case 7710
    frmMain.Show
    Exit Sub
Case 7740
    '// Right Click
    Exit Sub
Case 7755
    '// Right Click
    Exit Sub
Case Else
    '// Returned Mousemove events
    Exit Sub
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If fso.FileExists(App.Path & "\~Close.tmp") = True Then
    Kill App.Path & "\~Close.tmp"
    RemoveIcon
    End
End If
If UnloadMode = 2 Then
   ShutdownServer
   Exit Sub
End If
Cancel = 1
If 0 = 0 Then
Dim NoShowStatus As String
Dim NoShowMode As String
NoShowStatus = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XCyteServer\Settings", "NOSHOWDIALOG")
If NoShowStatus = "" Then NoShowStatus = 0
NoShowMode = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XCyteServer\Settings", "NOSHOWMODE")
If NoShowStatus = 1 Then
    Select Case NoShowMode
    Case 0
        RemoveIcon
        AddIcon "XCyteMail mail server"
        Me.Hide
        Exit Sub
    Case 1
        ShutdownServer
        Exit Sub
    Case Else
        RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "NOSHOWDIALOG", 0
        RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "NOSHOWMODE", "<INVALID>"
    End Select
End If
frmCloseInfo.Show
Else
Question = MsgBox("You are about to terminate the mailserver!" & vbCrLf & vbCrLf & "Click YES to proceed, or NO to minimise.", vbExclamation + vbYesNoCancel, "Mailserver termination")
Select Case Question
Case vbYes
ShutdownServer
Case vbNo
    frmMain.WindowState = 1
Case vbCancel
    ' Do Nothing
End Select
End If
End Sub

Private Sub Form_Resize()
If SysTray = True Then
If frmMain.WindowState = 1 Then frmMain.WindowState = 0: Me.Hide
End If
End Sub

Private Sub http_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If ControlLogging = True Then AppendLog "HTTP Connect request"
On Error GoTo clientGone
tryagain:
    For a = 1 To http.UBound
        If http(a).State = 0 Or http(a).State = 8 Then
            http(a).Close
            http(a).accept requestID
            Exit Sub
        End If
        DoEvents
    Next a
    DoEvents
    Num = http.UBound + 1
    Load http(Num)
    http(Num).accept requestID
    Exit Sub
clientGone:
AppendLog "   >HTTP Error: " & Err.Description
http(Index).Close
End Sub

Private Sub http_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
    If ControlLogging = True Then AppendLog "Data"
    If docroot = "\" Then docroot = App.Path
    
    Dim a As String, FileData As String, headers As Dictionary
    If http(Index).State <> 7 Then http(Index).Close: Exit Sub
    http(Index).getdata a
'    MsgBox a
    If bytesTotal = 8192 Or InStr(1, a, "multipart/form-data", TextCompare) Then
        http(Index).Tag = http(Index).Tag & a
        stats.List(Index) = "Incoming File..."
        AppendLog "   >Incoming file from: " & http(Index).RemoteHost & "(" & http(Index).RemoteHostIP & ")"
        Exit Sub
    Else
        a = http(Index).Tag & a
        http(Index).Tag = ""
    End If
    
    If a = "" Then
        http(Index).Close
        Exit Sub
    End If
    otherheaders = Mid(a & vbNewLine, InStr(1, a, vbNewLine) + 2)
    otherheaders = Mid(otherheaders, 1, InStr(1, otherheaders, vbNewLine & vbNewLine) - 1)
    
    Set headers = parseheaders(CStr(otherheaders))
    
    If InStr(1, a, vbNewLine & vbNewLine) > 0 Then postdata = Mid(a, InStr(1, a, vbNewLine & vbNewLine) + 4) Else postdata = ""
    
    If CLng(headers("Content-length")) > Len(postdata) Then
        stats.List(Index) = "Awaiting POST data"
        http(Index).Tag = http(Index).Tag & a
        Exit Sub
    End If
    
    If headers("Content-type") = "application/x-www-form-urlencoded" And IsEmpty(headers("Content-length")) Then
        http(Index).Tag = http(Index).Tag & a
        Exit Sub
    End If
    a = Left(a, InStr(1, a, vbNewLine) - 1)
    a = Mid(a, InStr(1, a, " ") + 1)
    a = Left(a, InStr(1, a, " ") - 1)

    
    While Mid(a, 1, 3) = "/.."
        a = Mid(a, 4)
    Wend
    
    If Right(a, 1) = "/" Then a = a & Default
    
    'seperated the request string into filename and GET data
    If Not CBool(InStr(1, a, "?")) Then
        a = a & "?"
    End If
    cmd = Left(a, InStr(1, a, "?") - 1)
    Data = Mid(a, InStr(1, a, "?") + 1)
    cmd = Replace(cmd, "/", "\")
    cmd = Replace(cmd, "%20", " ")
    
    header = "HTTP/1.0 200 OK" & vbNewLine & "Server: Xcyte mailserver" & vbNewLine & "Host: " & _
    http(Index).LocalIP & vbNewLine & "Connection: close" & vbNewLine
    '// Perform a simple check on the command to see if the requested file exists
    '// First check for code generated pages
'    Stop
DoSel:
Dim i As String
'Stop

DoEvents
    If LCase(Right(cmd, 3)) = "jpg" Then
        '// Handle the image file routine
        Open fso.BuildPath(frmAdmin.Rootpath.Text, fso.GetFilename(cmd)) For Binary As #1
        i = Space(LOF(1))
        Get #1, , i
        Close #1
        back = "Content-type: image/png" & vbCrLf & vbCrLf & i
        GoTo skipsel
    End If
    
    If LCase(Right(cmd, 3)) = "css" Then
        '// Handle the image file routine
        Open fso.BuildPath(frmAdmin.Rootpath.Text, fso.GetFilename(cmd)) For Binary As #1
        i = Space(LOF(1))
        Get #1, , i
        Close #1
        back = "Content-type: css" & vbCrLf & vbCrLf & i
        GoTo skipsel
    End If
    
    If LCase(Right(cmd, 3)) = "gif" Then
        '// Handle the image file routine
        Open fso.BuildPath(frmAdmin.Rootpath.Text, fso.GetFilename(cmd)) For Binary As #1
        i = Space(LOF(1))
        Get #1, , i
        Close #1
        back = "Content-type: GIF" & vbCrLf & vbCrLf & i
        GoTo skipsel
    End If
    
    Debug.Print cmd
    Select Case LCase(fso.GetParentFolderName(cmd))
    Case "\cgi-bin\img"
        Open fso.BuildPath(frmAdmin.Rootpath.Text, fso.GetFilename(cmd)) For Binary As #1
        i = Space(LOF(1))
        Get #1, , i
        Close #1
        back = "Content-type: image/png" & vbCrLf & vbCrLf & i
    Case "\cgi-bin"
        back = dowebsite(fso.GetFilename(cmd), tophpvariables(CStr(Data), CStr(postdata), headers("cookie")), headers, header, Index)
        'GoTo skipsel
    Case Else
        header = ""
        back = "HTTP/1.0 302 FOUND" & vbNewLine & "Server: XCyteMail" & vbNewLine & "Host: " & _
    http(Index).LocalIP & vbNewLine & "Url: /cgi-bin/login.asp" & vbNewLine & "Location: /cgi-bin/login.asp" & vbNewLine & "Connection: close" & vbNewLine & _
    vbNewLine
    'GoTo UseNormalHeader
    End Select
skipsel:
        'header = ""
'UseNormalHeader:
    back = header & back
On Error GoTo getout
    
    While Len(back) > 0
        http(Index).SendData Mid(back, 1, frmAdmin.MaxBuf.Text)
        back = Mid(back, frmAdmin.MaxBuf.Text + 1)
'        t = Timer + 0.1
'        While t > Timer
            DoEvents
'        Wend
    Wend
outofhere:
 '   t = Timer + 0.1
 '   While t > Timer
        DoEvents
 '   Wend
    
    On Error Resume Next
getout:
    http(Index).Close
    
End Sub

Public Sub Popin()
Do Until SideBar.Left <= 0 - SideBar.Width + Popper.Width + 170
    SideBar.Left = SideBar.Left - 100
    Sleep 1
    DoEvents
    SvrLog.Left = SideBar.Left + SideBar.Width
    SvrLog.Width = Me.Width - (SideBar.Width + SideBar.Left)
Loop
For x = 0 To SideObject.Count - 1
SideObject(x).Visible = False
Next x
Popper.Picture = ImageList1.ListImages(1).Picture
'Frame1.Visible = False
SvrLog.Width = Me.Width - (SideBar.Width + SideBar.Left)
End Sub

Public Sub PopOut()
For x = 0 To SideObject.Count - 1
SideObject(x).Visible = True
Next x
Popper.Picture = ImageList1.ListImages(2).Picture
'Frame1.Visible = True
Do Until SideBar.Left >= 0
    SideBar.Left = SideBar.Left + 100
    Sleep 1
    DoEvents
    SvrLog.Left = SideBar.Left + SideBar.Width
SvrLog.Width = Me.Width - (SideBar.Width + SideBar.Left)
Loop
SvrLog.Width = Me.Width - (SideBar.Width + SideBar.Left)
End Sub

Private Sub pop3_ConnectionRequest(ByVal requestID As Long)
    Dim a As POP3Session
    If ControlLogging = True Then
    AppendLog "POP3 Connection request"
    AppendLog ">"
    AppendLog "Creating POP3 Worker Thread"
    End If

    Set a = New POP3Session
    
    a.ws.accept requestID
    Randomize
    
    ts = "<" & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & ">"
    a.Timestamp = ts
    a.ws.SendData "+OK XCyte Mail Server " & ts & vbCrLf
    
    'a.Show
End Sub

Private Sub pop3_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    pop3.Close
    pop3.listen
End Sub

Private Sub POP3Check_Click()
On Error Resume Next
If POP3Check.value = 0 Then
  StopService ("POP3")
Else
   StartService ("POP3")
End If
Me.SetFocus
End Sub

Private Sub Popper_Click()
If Popper.Tag = 0 Then
   Popin
   Popper.Tag = 1
Else
   PopOut
   Popper.Tag = 0
End If
End Sub

Private Sub ShutdownTimer_Timer()
ShutdownCount = ShutdownCount - 1
If ShutdownCount = 0 Then
   ShutdownServer
End If
End Sub

Private Sub SideObject_Click(Index As Integer)
Select Case Index
Case 0
    ' Start/stop the server
    If ServerActive = False Then
        StartServer
    Else
        StopServer
    End If
Case 1
    ' Shutdown the server
    ShutdownServer
Case 2
    ' Flush the account inboxes
    FlushAccounts
Case 3
    ' Show the administrator panel
    frmAdmin.Show
Case 4
    ' Show the inboxes panel
    frmUsers.Show
Case 5
    ' Show the diagnostics panel
    frmDiag.Show
End Select
End Sub

Private Sub SMTPCheck_Click()
If Called = True Then Exit Sub
If SMTPCheck.value = 0 Then
  StopService ("SMTP")
Else
   StartService ("SMTP")
End If
End Sub

Private Sub SvrLog_KeyDown(KeyCode As Integer, Shift As Integer)
If Editing = True Then
   If KeyCode = vbKeyReturn Then
   AppendLog vbCrLf
   End If
End If
If KeyCode = 27 Then
   Shifted = False
   SvrLog.Text = Left(SvrLog.Text, InStrRev(SvrLog.Text, ":", Len(SvrLog.Text), vbTextCompare) - 1)
End If
If Editing = True Then
   If Shift = 1 And KeyCode = 186 Then
      AppendLog vbCrLf & ":"
      Shifted = True
      SvrLog.Text = Left(SvrLog.Text, Len(SvrLog.Text) - 1)
   End If
End If

If KeyCode = vbKeyDelete Then KeyCode = 0
SvrLog.SelStart = Len(SvrLog.Text)
End Sub

Private Sub SvrLog_KeyPress(KeyAscii As Integer)
'Stop
On Error Resume Next
If Right(SvrLog.Text, 1) = ">" Then
If KeyAscii = 8 Then KeyAscii = 0
End If

'If Editing = True Then
'If Right(SvrLog.Text, 1) = ":" Then
'If KeyAscii = 8 Then KeyAscii = 0
'End If
'End If

If KeyAscii = 13 Then
   Processcommand Right(SvrLog.Text, Len(SvrLog.Text) - InStrRev(SvrLog.Text, vbCrLf, -1, vbTextCompare) - 2)
   KeyAscii = 0
End If
'SvrLog.Text = SvrLog.Text & vbCrLf & ">"
SvrLog.SelStart = Len(SvrLog.Text)
End Sub

Private Sub SvrLog_KeyUp(KeyCode As Integer, Shift As Integer)
SvrLog.SelStart = Len(SvrLog.Text)
End Sub

Private Sub Timer2_Timer()
'SvrLog.SelStart = Len(SvrLog.Text)
End Sub

Private Sub Timer3_Timer()
If Autopop = False Then
   SideBar.Left = 0
   Timer3.Enabled = False
   Exit Sub
End If
If GetHandle = SideBar.hWnd Then
   PopOut
Else
   expTimer1.Enabled = True
End If
End Sub

Private Sub Webmailcheck_Click()
If Webmailcheck.value = 0 Then
  StopService ("WEBMAIL")
Else
   StartService ("WEBMAIL")
End If
End Sub


Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If ControlLogging = True Then AppendLog "Ws connection request"
    For a = 1 To ws.UBound
        If ws(a).State = 0 Or ws(a).State = 8 Or ws(a).State = 9 Then Exit For
    Next
    On Error Resume Next
    Load ws(a)
    On Error GoTo 0
    ws(a).Close
    ws(a).accept requestID
    ws(a).SendData "220 XCyteMail SMTP Server" & vbCrLf
    Set c = New inmail
    Set obj(a) = c
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim a As String
    On Error Resume Next
    ws(Index).getdata a
    obj(Index).moreincomming a
    obj(Index).parsebuffer
    c = obj(Index).outbuffer
    If obj(Index).ToggleClose = True Then
        ws(Index).Close
        Exit Sub
    End If
    If Len(c) > 0 Then ws(Index).SendData c: obj(Index).outbuffer = ""
    If Mid(c, 1, 3) = "221" Then ws(Index).Close
End Sub

Private Sub ws_Error(Index As Integer, ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ws(Index).Close
    If Index = 0 Then ws(0).listen
End Sub

Public Sub Timer1_Timer()
'Stop
    If ControlLogging = True Then AppendLog "Checked mail que..."
'    Stop
    If ws(ws.UBound).State = 8 And ws.UBound > 0 Then Unload ws(ws.UBound)
    If http(http.UBound).State = 8 And http.UBound > 0 Then Unload http(http.UBound)
    
    ws(0).Close
    ws(0).listen
        
    On Error Resume Next
    http(0).Close
    http(0).listen
        
    On Error GoTo 0
    pop3.Close
    pop3.listen
    Dim s As SendEmail
    DoEvents
    Do Until SendEmail.MailPending = False
    DoEvents
    Loop
    For a = 0 To frmMain.File1.ListCount - 1
        If Mid(frmMain.File1.List(a), 1, 1) <> "!" Then
            Set s = New SendEmail
            s.setup fso.BuildPath(frmMain.File1.Path, frmMain.File1.List(a))
            sentemails.Add s
        End If
    Next a
    
    If frmMain.File1.ListCount = 0 And sentemails.Count > 0 Then
        Set sentemails = New Collection
    End If
    
    frmMain.File1.Refresh
    'AppendLog ">"
End Sub
