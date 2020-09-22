VERSION 5.00
Begin VB.Form frmInit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Initializing installation..."
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "frmInit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6270
      Top             =   2940
   End
   Begin XCyteMail.Line3D Line3D1 
      Height          =   375
      Left            =   -450
      TabIndex        =   4
      Top             =   870
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   661
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   7050
      TabIndex        =   0
      Top             =   0
      Width           =   7050
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   240
         Picture         =   "frmInit.frx":000C
         ScaleHeight     =   705
         ScaleWidth      =   720
         TabIndex        =   1
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while the application is configured..."
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
         Left            =   1200
         TabIndex        =   3
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Initializing XCyteMail installation"
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
         Left            =   1200
         TabIndex        =   2
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   495
      Shape           =   3  'Circle
      Top             =   2445
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "Creating user accounts database..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   705
      TabIndex        =   8
      Top             =   2400
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   495
      Shape           =   3  'Circle
      Top             =   2130
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Creating default configuration scripts..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   705
      TabIndex        =   7
      Top             =   2085
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   495
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "Performing the following actions..."
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1395
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "Creating and verifying application directories..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1755
      Width           =   6615
   End
End
Attribute VB_Name = "frmInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SetupOK As Boolean

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   Open App.Path & "\XCyte.Bin" For Output As #1
   ' Generate an encryption code and write 2 file
   Dim Gencode As String
   Dim RandomString
   Randomize
   For x = 0 To 128
   RandomString = Chr(Int(Rnd * 128))
   RandomString = (RandomString) & "00" & (Asc(RandomString) ^ 2)
   RandomString = Mid(RandomString, Int(Rnd * (Len(RandomString) - 1) + 1), 1)
   Gencode = Gencode & RandomString
   Next x
   Gencode = "{Public Key}" & vbCrLf & "XProto" & Gencode
   Print #1, Gencode
   
   Close #1
   RunTimeCom.RuntimeStartup
   SetupOK = True
   Me.Hide
End Sub
