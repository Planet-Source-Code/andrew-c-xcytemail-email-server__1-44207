VERSION 5.00
Begin VB.Form frmDiag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server diagnostics"
   ClientHeight    =   3780
   ClientLeft      =   3510
   ClientTop       =   2100
   ClientWidth     =   6375
   Icon            =   "frmDiag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6375
   Begin VB.CommandButton Command2 
      Caption         =   "Execute Timer"
      Height          =   495
      Left            =   1815
      TabIndex        =   10
      Top             =   3045
      Width           =   1215
   End
   Begin XCyteMail.MX MX1 
      Left            =   3630
      Top             =   3120
      _ExtentX        =   900
      _ExtentY        =   741
   End
   Begin VB.Frame Frame2 
      Caption         =   "MX/DNS:"
      Height          =   3735
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   3015
      Begin VB.ComboBox DNSList 
         Height          =   315
         Left            =   1515
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   165
         Width           =   1380
      End
      Begin VB.CommandButton Command1 
         Caption         =   "MX Lookup"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "MX Lookup:"
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2775
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Text            =   "[None]"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Text            =   "http://www.yahoo.com"
            Top             =   480
            Width           =   2535
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   1080
            Picture         =   "frmDiag.frx":0442
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Returned MX:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Hostname:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Label UseMX 
         Caption         =   "Using DNS server:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Outgoing mail que:"
      Height          =   3750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3225
      Begin VB.FileListBox File1 
         Height          =   3405
         Left            =   105
         Pattern         =   "*.txt"
         TabIndex        =   1
         Top             =   225
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmDiag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MX_Table As New Dictionary

Private Sub Command1_Click()
'On Error Resume Next
Text2.Text = mxlookup(Text1.Text)
End Sub

Private Sub Command2_Click()
Call frmMain.Timer1_Timer
End Sub

