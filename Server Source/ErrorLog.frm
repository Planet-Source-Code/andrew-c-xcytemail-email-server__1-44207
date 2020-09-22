VERSION 5.00
Begin VB.Form ErrorLog 
   Caption         =   "Extended error log"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   Icon            =   "ErrorLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
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
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "ErrorLog.frx":0442
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "ErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub WriteError(sString)
Text1.Text = Text1.Text & vbCrLf & sString
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Form_Activate()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Form_Load()
Text1.SelStart = Len(Text1.Text)
End Sub
