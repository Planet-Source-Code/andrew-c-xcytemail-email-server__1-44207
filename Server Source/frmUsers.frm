VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User manager"
   ClientHeight    =   4455
   ClientLeft      =   2235
   ClientTop       =   3525
   ClientWidth     =   8415
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8415
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   7680
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   9120
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView UserList 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7858
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Mailbox"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Message Count"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuSubAccount 
      Caption         =   "&File"
      Begin VB.Menu mnuSubNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuSubOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubExit 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub RefreshUsers()
UserList.ListItems.Clear

strProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & GetWindir & "\xmbox\Accounts.mdb;"
Dim objconn As ADODB.Connection
Dim objrs As ADODB.Recordset
Set objconn = New ADODB.Connection
Set objrs = New ADODB.Recordset

objconn.Open strProvider
strQuery = "SELECT * FROM Logins"
strQuery = strQuery & " ORDER BY " & "Username" & " ASC"
Set objrs = objconn.Execute(strQuery)

AdderIndex = 0
'objrs.MoveFirst
Do While Not objrs.EOF
UserList.ListItems.Add , , objrs(AdderIndex)
objrs.MoveNext
Loop

End Sub

Public Function GetInboxListing(inbox As String)
File1.Path = inbox
File1.Refresh
GetInboxListing = File1.ListCount - 1
End Function

Private Sub Form_Activate()
RefreshUsers
End Sub

Private Sub mnuSubNew_Click()
frmNewUser.txtusername.Text = ""
frmNewUser.txtPassword.Text = ""
frmNewUser.altEmail.Text = ""
frmNewUser.smsemail.Text = ""
frmNewUser.Show
End Sub

Private Sub UserList_DblClick()
' Open the selected account
Dim newwindow As New AccountView
newwindow.Show
newwindow.Caption = "Account Viewer: " & UserList.SelectedItem.Text
newwindow.File1.Path = subfolder(UserList.SelectedItem.Text)
newwindow.File1.Refresh
For x = 0 To newwindow.File1.ListCount - 1
If LCase(newwindow.File1.List(x)) = "!account.txt" Then GoTo skipper
Open newwindow.File1.Path & "\" & newwindow.File1.List(x) For Input As #1
Dim sFrom As String
Dim sTo As String
Dim NullBuffer As String
Dim sDate As String
Input #1, sFrom
Input #1, sTo
Input #1, NullBuffer
Input #1, sDate
Close #1
Set adder = newwindow.List1.ListItems.Add(, , sDate)
adder.SubItems(1) = sFrom
adder.SubItems(2) = sTo
adder.Tag = newwindow.File1.Path & "\" & newwindow.File1.List(x)
skipper:
Next x
End Sub
