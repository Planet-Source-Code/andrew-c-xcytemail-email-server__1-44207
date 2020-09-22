VERSION 5.00
Begin VB.Form frmCloseInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Note"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmCloseInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin XCyteMail.Line3D Line3D1 
      Height          =   255
      Left            =   -90
      TabIndex        =   7
      Top             =   885
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&No"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Don't ask this question again"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Do you want to send it to the system-tray instead?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "You are about to close the mailserver application. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   240
         Picture         =   "frmCloseInfo.frx":000C
         Top             =   120
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmCloseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaveMode As String
Private Sub Command1_Click()
' Minimise to systray
SaveMode = 0
SaveQuestion
RemoveIcon
AddIcon "XCyteMail mail server"
Me.Hide
frmMain.Hide
End Sub

Private Sub Command2_Click()
SaveMode = 1
SaveQuestion
ShutdownServer
End Sub

Private Sub SaveQuestion()
' If the user clicked the button, save the preference
If Check1.value = 1 Then
   RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "NOSHOWDIALOG", 1
   RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "NOSHOWMODE", SaveMode
End If
End Sub

Private Sub Command3_Click()
Me.Hide
Check1.value = 0
End Sub

