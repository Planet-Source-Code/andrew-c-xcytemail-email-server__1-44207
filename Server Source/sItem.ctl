VERSION 5.00
Begin VB.UserControl sItem 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7395
   ScaleHeight     =   1080
   ScaleWidth      =   7395
   Begin VB.PictureBox sArray 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   0
      Left            =   0
      ScaleHeight     =   1080
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin VB.CommandButton sStart 
         Caption         =   "Start"
         Height          =   285
         Index           =   0
         Left            =   6525
         TabIndex        =   2
         Top             =   675
         Width           =   675
      End
      Begin VB.CommandButton sUNINST 
         Caption         =   "Uninstall"
         Height          =   285
         Index           =   0
         Left            =   5685
         TabIndex        =   1
         Top             =   675
         Width           =   795
      End
      Begin VB.Label servName 
         BackStyle       =   0  'Transparent
         Caption         =   "[ServiceName]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2070
      End
      Begin VB.Label sDescriptor 
         BackStyle       =   0  'Transparent
         Caption         =   "[Description]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   675
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   255
         Width           =   4695
      End
      Begin VB.Label sType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Component Type: [Type]"
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   0
         Left            =   5160
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label sAuth 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Service Author: [Author]"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   5160
         TabIndex        =   3
         Top             =   285
         Width           =   2025
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   -30
         Top             =   -30
         Width           =   7380
      End
   End
End
Attribute VB_Name = "sItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim CompressedState As Boolean

Public Property Let Compressed(ByVal State As Boolean)
' This is a state
CompressedState = State
If CompressedState = True Then
    CompressItem
Else
    DecompressItem
End If
End Property

Public Property Get Compressed() As Boolean
Compressed = CompressedState
End Property

Public Property Let ServiceName(ByVal sString As String)
servName(0).Caption = sString
UserControl.Tag = sString
End Property

Public Property Get ServiceName() As String
ServiceName = servName(0).Caption
End Property

Public Property Let Description(ByVal sString As String)
sDescriptor(0).Caption = sString
End Property

Public Property Get Description() As String
ServiceName = sDescriptor(0).Caption
End Property

Public Property Let ServiceType(ByVal sString As String)
sType(0).Caption = sString
End Property

Public Property Get ServiceType() As String
ServiceType = sType(0).Caption
End Property

Public Property Let ServiceAuthor(ByVal sString As String)
sAuth(0).Caption = sString
End Property

Public Property Get ServiceAuthor() As String
ServiceAuthor = sAuth(0).Caption
End Property

Private Sub Image1_Click()
DecompressItem
End Sub

Private Sub CompressItem()
UserControl.Height = 330
sType(0).Visible = False
sAuth(0).Visible = False
sArray(0).BackColor = &H800000
servName(0).ForeColor = &HFFFFFF
sDescriptor(0).Visible = False
End Sub

Private Sub DecompressItem()
Dim x As Long
Dim GoFlag As Boolean

For x = 0 To frmAdmin.sItem.Count - 1
frmAdmin.sItem(x).Compressed = True
If x <> 0 Then
frmAdmin.sItem(x).Top = frmAdmin.sItem(x - 1).Top + frmAdmin.sItem(x - 1).Height
End If
Next x
UserControl.Height = 1080
sType(0).Visible = True
sAuth(0).Visible = True
sDescriptor(0).Visible = True
sArray(0).BackColor = &HFFFFFF
servName(0).ForeColor = &H0&
sDescriptor(0).ForeColor = &H0&
sType(0).ForeColor = &H0&
sAuth(0).ForeColor = &H0&

' Compress all items again

' Move down all other items
For x = 0 To frmAdmin.sItem.Count - 1
If frmAdmin.sItem(x).ServiceName = servName(0).Caption Then
    GoFlag = True
Else
    If GoFlag = True Then
        frmAdmin.sItem(x).Top = frmAdmin.sItem(x - 1).Top + frmAdmin.sItem(x - 1).Height
    End If
End If
Next x
End Sub

Private Sub servName_Click(Index As Integer)
DecompressItem
End Sub

Private Sub UserControl_Initialize()
CompressItem
End Sub
