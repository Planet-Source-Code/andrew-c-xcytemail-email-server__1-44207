VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AccountView 
   Caption         =   "Account Viewer"
   ClientHeight    =   5265
   ClientLeft      =   2340
   ClientTop       =   2850
   ClientWidth     =   9180
   Icon            =   "AccountView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView List1 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7646
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date/Time"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "To"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   5040
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "AccountView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
List1.Left = 0
List1.Top = 0
List1.Width = Me.Width
List1.Height = Me.Height
End Sub

Private Sub List1_DblClick()
On Error Resume Next
Shell "notepad.exe " & List1.SelectedItem.Tag, vbNormalFocus
End Sub
