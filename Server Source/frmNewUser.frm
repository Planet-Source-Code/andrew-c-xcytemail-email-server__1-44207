VERSION 5.00
Begin VB.Form frmNewUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create new mailbox"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "C&ancel"
      Height          =   495
      Left            =   3270
      TabIndex        =   11
      Top             =   2955
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Create"
      Default         =   -1  'True
      Height          =   495
      Left            =   3270
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1245
      Left            =   105
      TabIndex        =   5
      Top             =   990
      Width           =   4620
      Begin VB.TextBox smsemail 
         Height          =   285
         Left            =   1395
         TabIndex        =   9
         Top             =   660
         Width           =   2910
      End
      Begin VB.TextBox altEmail 
         Height          =   285
         Left            =   1395
         TabIndex        =   7
         Top             =   225
         Width           =   2910
      End
      Begin VB.Label Label4 
         Caption         =   "SMS email:"
         Height          =   210
         Left            =   510
         TabIndex        =   8
         Top             =   675
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Alternative email:"
         Height          =   210
         Left            =   105
         TabIndex        =   6
         Top             =   240
         Width           =   4470
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   885
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   3525
   End
   Begin VB.TextBox txtusername 
      Height          =   285
      Left            =   885
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   180
      Left            =   75
      TabIndex        =   3
      Top             =   615
      Width           =   3060
   End
   Begin VB.Label lblhostname 
      Caption         =   "@hostname.com"
      Height          =   255
      Left            =   3135
      TabIndex        =   2
      Top             =   165
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Mailbox:"
      Height          =   255
      Left            =   210
      TabIndex        =   0
      Top             =   165
      Width           =   3735
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If fso.FolderExists(App.Path & "\email" & txtusername.Text) = True Then
   errdesc = "The mailbox name already exists!" & vbCrLf & "Choose another name"
   GoTo ErrorOccured
End If
If Len(txtusername.Text) < 3 Then
   errdesc = "The mailbox name is too short!" & vbCrLf & "The mailbox name must be longer than 3 characters."
   GoTo ErrorOccured
End If
If Len(txtPassword.Text) < 6 Then
   errdesc = "The password is too short!" & vbCrLf & "The password must be greater than 6 characters"
   GoTo ErrorOccured
End If

CreateAccount txtusername.Text, txtPassword.Text, altEmail.Text, smsemail.Text

'subfolder txtusername.Text
'Open subfolder(txtusername.Text) & "\!account.txt" For Output As #1
'Print #1, "pw: " & LCase(txtPassword.Text)
'Print #1, "alt: " & LCase(AltEmail.Text)
'Print #1, "sms: " & LCase(SMSEmail.Text)
'Close #1
MsgBox "Account " & txtusername.Text & " successfully created!", vbInformation, "Success!"
Me.Hide
frmUsers.RefreshUsers
Exit Sub
ErrorOccured:
MsgBox "Unable to create account!" & vbCrLf & vbCrLf & "Reason: " & errdesc, vbCritical, "Operation Failed"
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

