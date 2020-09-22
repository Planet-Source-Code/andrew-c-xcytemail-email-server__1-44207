VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Administration"
   ClientHeight    =   4095
   ClientLeft      =   2220
   ClientTop       =   3945
   ClientWidth     =   8205
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Obsolete"
      Height          =   3495
      Left            =   1680
      TabIndex        =   20
      Top             =   4800
      Width           =   5175
      Begin VB.TextBox MaxBuf 
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Text            =   "256"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox CookieTimeout 
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Text            =   "30"
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label9 
         Caption         =   "Cookie:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Buffer:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Login"
      TabPicture(0)   =   "frmAdmin.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "pMinLength"
      Tab(0).Control(1)=   "MinScroll"
      Tab(0).Control(2)=   "MinScroll1"
      Tab(0).Control(3)=   "uMinLength"
      Tab(0).Control(4)=   "accMin"
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(8)=   "Label4"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Networking"
      TabPicture(1)   =   "frmAdmin.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(4)=   "LocalHostname"
      Tab(1).Control(5)=   "httpport"
      Tab(1).Control(6)=   "AutoDNS"
      Tab(1).Control(7)=   "DNSSet1"
      Tab(1).Control(8)=   "DNSSet2"
      Tab(1).Control(9)=   "DNSSet3"
      Tab(1).Control(10)=   "DNSSet4"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Email"
      TabPicture(2)   =   "frmAdmin.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Mailforward"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Filters"
      TabPicture(3)   =   "frmAdmin.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(1)=   "List1"
      Tab(3).Control(2)=   "Check1"
      Tab(3).Control(3)=   "Command4"
      Tab(3).Control(4)=   "Command5"
      Tab(3).Control(5)=   "Command6"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Stubs"
      TabPicture(4)   =   "frmAdmin.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label13"
      Tab(4).Control(1)=   "Command2"
      Tab(4).Control(2)=   "List3"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Webmail"
      TabPicture(5)   =   "frmAdmin.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command3"
      Tab(5).Control(1)=   "Rootpath"
      Tab(5).Control(2)=   "Label14"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Services"
      TabPicture(6)   =   "frmAdmin.frx":04EA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Command7"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Picture1"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.PictureBox Picture1 
         Height          =   1830
         Left            =   210
         ScaleHeight     =   1770
         ScaleWidth      =   7680
         TabIndex        =   41
         Top             =   510
         Width           =   7740
         Begin VB.VScrollBar VScroll1 
            Height          =   1770
            Left            =   7410
            TabIndex        =   43
            Top             =   0
            Width           =   270
         End
         Begin XCyteMail.sItem sItem 
            Height          =   330
            Index           =   0
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   582
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Populate"
         Height          =   495
         Left            =   5400
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Mailforward 
         Caption         =   "Enable mail forwarding"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox DNSSet4 
         Height          =   285
         Left            =   -71280
         MaxLength       =   3
         TabIndex        =   37
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox DNSSet3 
         Height          =   285
         Left            =   -71880
         MaxLength       =   3
         TabIndex        =   36
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox DNSSet2 
         Height          =   285
         Left            =   -72480
         MaxLength       =   3
         TabIndex        =   35
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox DNSSet1 
         Height          =   285
         Left            =   -73080
         MaxLength       =   3
         TabIndex        =   34
         Top             =   2640
         Width           =   375
      End
      Begin VB.CheckBox AutoDNS 
         Caption         =   "Automatically locate the DNS server"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   2280
         Width           =   3855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -69075
         TabIndex        =   32
         Top             =   615
         Width           =   735
      End
      Begin VB.TextBox Rootpath 
         Height          =   285
         Left            =   -73800
         TabIndex        =   31
         Top             =   600
         Width           =   4695
      End
      Begin VB.ListBox List3 
         Height          =   2400
         Left            =   -74880
         TabIndex        =   29
         Top             =   720
         Width           =   6255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit Stub"
         Height          =   615
         Left            =   -68520
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox httpport 
         Height          =   285
         Left            =   -72480
         TabIndex        =   25
         Text            =   "80"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Add Rule"
         Height          =   495
         Left            =   -68280
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Remove Rule"
         Height          =   495
         Left            =   -68280
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete All"
         Height          =   495
         Left            =   -68280
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable mail filter"
         Height          =   255
         Left            =   -74640
         TabIndex        =   15
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox LocalHostname 
         Height          =   285
         Left            =   -72480
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox pMinLength 
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "6"
         Top             =   960
         Width           =   375
      End
      Begin VB.VScrollBar MinScroll 
         Height          =   285
         Left            =   -72840
         Max             =   10
         Min             =   3
         TabIndex        =   5
         Top             =   945
         Value           =   3
         Width           =   255
      End
      Begin VB.VScrollBar MinScroll1 
         Height          =   285
         Left            =   -72840
         Max             =   10
         Min             =   3
         TabIndex        =   4
         Top             =   600
         Value           =   3
         Width           =   255
      End
      Begin VB.TextBox uMinLength 
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "3"
         Top             =   615
         Width           =   375
      End
      Begin VB.TextBox accMin 
         Height          =   285
         Left            =   -73200
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin MSComctlLib.ListView List1 
         Height          =   2535
         Left            =   -74520
         TabIndex        =   16
         Top             =   1320
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Rule"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Use DNS Server:                 :            :             :"
         Height          =   255
         Left            =   -74625
         TabIndex        =   38
         Top             =   2655
         Width           =   4455
      End
      Begin VB.Label Label14 
         Caption         =   "WWW Root:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label13 
         Caption         =   "Select STUB for editing:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label Label12 
         Caption         =   "Operate HTTP server on port:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label Label7 
         Caption         =   $"frmAdmin.frx":0506
         Height          =   615
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   7575
      End
      Begin VB.Label Label11 
         Caption         =   $"frmAdmin.frx":0599
         Height          =   615
         Left            =   -74760
         TabIndex        =   13
         Top             =   480
         Width           =   7575
      End
      Begin VB.Label Label5 
         Caption         =   "Hostname:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Min password length:                characters."
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   975
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Min username length:                characters."
         Height          =   255
         Left            =   -74775
         TabIndex        =   9
         Top             =   630
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Max account size:"
         Height          =   255
         Left            =   -74565
         TabIndex        =   8
         Top             =   1455
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "bytes."
         Height          =   255
         Left            =   -71550
         TabIndex        =   7
         Top             =   1470
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserDNS As String
Public OldVal As Integer
Public sLoaded As Boolean

Private Sub AutoDNS_Click()
If AutoDNS.value = 1 Then
   DNSSet1.Enabled = False
   DNSSet2.Enabled = False
   DNSSet3.Enabled = False
   DNSSet4.Enabled = False
Else
   DNSSet1.Enabled = True
   DNSSet2.Enabled = True
   DNSSet3.Enabled = True
   DNSSet4.Enabled = True
End If
End Sub

Private Sub Check2_Click()

End Sub

Private Sub Command1_Click()
If Rootpath.Text = "" Or Rootpath.Text = " " Then Rootpath.Text = App.Path & "\www-mail"
    If fso.FolderExists(Rootpath.Text) = False Then
    If MsgBox("The specified webmail resource directory does not exist." & vbCrLf & "Do you want to create it?", vbQuestion + vbYesNo, "Webmail") = vbYes Then
        On Error GoTo CannotCreate
        MkDir Rootpath.Text
        GoTo DoneOK
CannotCreate:
        MsgBox "An error occured while creating the specified webmail directory.", vbCritical, "Operation failed"
    End If
    End If
DoneOK:
UserDNS = DNSSet1.Text & "." & DNSSet2.Text & "." & DNSSet3.Text & "." & DNSSet4.Text
HostName = LocalHostname.Text
Me.Hide
SaveSettings
End Sub

Private Sub Command2_Click()
frmFilters.Show
End Sub

Private Sub Command3_Click()
Rootpath.Text = Browse_folder.Browse_folder("C:\")
If fso.FileExists(Rootpath.Text & "\header.htm") = False And fso.FileExists(Rootpath.Text & "\header.html") = False Then
    MsgBox "The directory is invalid." & vbCrLf & "header.html not found", vbExclamation, "Cannot use selected directory"
End If
End Sub

Private Sub Command4_Click()
Open App.Path & "\Config\filters.cfg" For Output As #1
Close #1
LoadRules
End Sub

Private Sub Command5_Click()
Dim sComp As String
Dim RuleCount As Integer
RuleCount = 1
sComp = Left(List1.SelectedItem.Text, 6)
' Remove the selected rule
Open App.Path & "\Config\filters.cfg" For Input As #1
Open App.Path & "\Config\filters.cf~" For Output As #2
Do Until EOF(1)
Input #1, Void
Input #1, RuleName
Input #1, RuleType
Input #1, RuleCFG
Input #1, Void
If LCase(sComp) = LCase(RuleName) Then
   ' Omit the rule
   GoTo SkipRule
Else
   NewRuleName = "Rule" & RuleCount & ":"
   Print #2, "{"
   Print #2, NewRuleName
   Print #2, RuleType
   Print #2, RuleCFG
   Print #2, "}"
End If
SkipRule:
RuleCount = RuleCount + 1
Loop
Close #2
Close #1
Kill App.Path & "\Config\Filters.cfg"
Name App.Path & "\Config\Filters.cf~" As App.Path & "\Config\Filters.cfg"
LoadRules
End Sub

Private Sub Command6_Click()
frmAddRule.Option1.value = True
frmAddRule.fKeyword.Text = ""
frmAddRule.fSubject.Text = ""
frmAddRule.fSender.Text = ""
frmAddRule.Show
End Sub

Private Sub Command7_Click()
PopulateServices
End Sub



Private Sub DNSSet1_Change()
On Error GoTo SkipFunct
If Len(DNSSet1.Text) = 3 Then
   DNSSet2.SetFocus
   DNSSet2.SelStart = 0
   DNSSet2.SelLength = Len(DNSSet2.Text)
End If
SkipFunct:
End Sub

Private Sub DNSSet1_GotFocus()
DNSSet1.SelStart = 0
DNSSet1.SelLength = Len(DNSSet1.Text)
End Sub

Private Sub DNSSet2_Change()
On Error GoTo SkipFunct
If Len(DNSSet2.Text) = 3 Then
   DNSSet3.SetFocus
   DNSSet3.SelStart = 0
   DNSSet3.SelLength = Len(DNSSet3.Text)
End If
SkipFunct:
End Sub

Private Sub DNSSet2_GotFocus()
DNSSet2.SelStart = 0
DNSSet2.SelLength = Len(DNSSet2.Text)
End Sub

Private Sub DNSSet3_Change()
On Error GoTo SkipFunct
If Len(DNSSet3.Text) = 3 Then
   DNSSet4.SetFocus
   DNSSet4.SelStart = 0
   DNSSet4.SelLength = Len(DNSSet4.Text)
End If
SkipFunct:
End Sub

Private Sub DNSSet3_GotFocus()
DNSSet3.SelStart = 0
DNSSet3.SelLength = Len(DNSSet3.Text)
End Sub

Private Sub DNSSet4_GotFocus()
DNSSet4.SelStart = 0
DNSSet4.SelLength = Len(DNSSet4.Text)
End Sub

Private Sub Form_Activate()
If sLoaded = False Then
    PopulateServices
End If
End Sub

Private Sub httpport_Change()
If IsNumeric(httport) = False Then
   MsgBox "The HTTP port must be a numeric digit between 1 and 65535!" & vbCrLf & "Only enter numbers please.", vbExclamation, "Bad port format"
End If
httpport.Text = ""
httpport.SetFocus
End Sub

Private Sub Mailforward_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Mailforward.value = 1 Then
    If MsgBox("Please note that with mail forwarding enabled, a malicious party may 'fake' email through your mail server." & vbCrLf & "Are you sure you want to enable mail forwarding?", vbQuestion + vbOKCancel, "Security warning") = vbOK Then
        '// Do nothing
    Else
        Mailforward.value = 0
    End If
End If
End Sub

Private Sub MinScroll_Change()
pMinLength.Text = MinScroll.value
End Sub

Private Sub MinScroll1_Change()
uMinLength.Text = MinScroll1.value
End Sub

Public Sub LoadRules()
eRuleCount = 0
If FileExists(App.Path & "\Config\Filters.cfg") = False Then Exit Sub
Dim RuleName As String
Dim RuleType As String
Dim RuleCFG As String
List1.ListItems.Clear
' Load the rules out of the file
Open App.Path & "\Config\Filters.cfg" For Input As #1
Do Until EOF(1)
Line Input #1, Void 'First header block
Line Input #1, RuleName
Line Input #1, RuleType
Line Input #1, RuleCFG
Line Input #1, Void 'Tail header block
ProcessRule RuleName, RuleType, RuleCFG
eRuleCount = eRuleCount + 1
Loop
Close #1
End Sub

Private Sub ProcessRule(sRuleName As String, sRuleType As String, sRuleCFG As String)
Dim RuleString
sRuleType = LCase(Right(sRuleType, Len(sRuleType) - InStr(1, sRuleType, ":", vbTextCompare - 1)))
sRuleType = Right(sRuleType, Len(sRuleType) - 1)
sRuleCFG = LCase(Right(sRuleCFG, Len(sRuleCFG) - Len(sRuleType) - 2))
Select Case sRuleType
Case "keyword"
RuleString = sRuleName & " Filter out all messages that contain the keywords '" & sRuleCFG & "'"
Case "sender"
RuleString = sRuleName & " Filter messages that come from '" & sRuleCFG & "'"
Case "subject"
RuleString = sRuleName & " Filter messages with '" & sRuleCFG & "' in the subject line"
Case "recipient"
RuleString = sRuleName & " Filter messages to '" & sRuleCFG & "'"
Case Else
RuleString = "Malformed rule format"
End Select
List1.ListItems.Add , , RuleString
End Sub

Private Sub VScroll1_Change()
MsgBox "Scrollbar not done yet"
End Sub

Private Sub PopulateServices()
' Create the service list
On Error GoTo enderror
Dim ServiceList() As String
Dim FileHandle0 As Long
Dim sString As String
Dim ServiceCount As Integer
ServiceCount = 0
FileHandle0 = FreeFile
Open App.Path & "\Config\Services.cfg" For Input As #FileHandle0
Do Until EOF(FileHandle0)
Input #FileHandle0, sString
If Left(sString, 1) <> ";" Then
ReDim Preserve ServiceList(ServiceCount)
ServiceList(ServiceCount) = sString
ServiceCount = ServiceCount + 1
End If
Loop
Close #FileHandle0

' Now create an object for each item
Dim sName As String
Dim sOwner As String
Dim sAuthor As String
Dim sDesc As String
Dim RawData As String
For x = 0 To UBound(ServiceList)
RawData = ServiceList(x)
sName = Left(RawData, InStr(1, RawData, "!", vbTextCompare) - 1)
RawData = Right(RawData, Len(RawData) - Len(sName) - 1)
sOwner = Left(RawData, InStr(1, RawData, "!", vbTextCompare) - 1)
RawData = Right(RawData, Len(RawData) - Len(sOwner) - 1)
sAuthor = Left(RawData, InStr(1, RawData, "!", vbTextCompare) - 1)
RawData = Right(RawData, Len(RawData) - Len(sAuthor) - 1)
sDesc = RawData

If x = 0 Then
    ' Just use the existing menu item
    sItem(0).ServiceAuthor = sAuthor
    sItem(0).ServiceName = sName
    sItem(0).ServiceType = sOwner
    sItem(0).Description = sDesc
Else
    Load sItem(x)
    sItem(x).Visible = True
    sItem(x).Top = sItem(x - 1).Top + sItem(x).Height
    sItem(x).ServiceAuthor = sAuthor
    sItem(x).ServiceName = sName
    sItem(x).ServiceType = sOwner
    sItem(x).Description = sDesc
End If
Next x
enderror:
End Sub
