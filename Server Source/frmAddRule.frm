VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddRule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add filter rule"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmAddRule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.txt"
      DialogTitle     =   "Import sender black list"
      Filter          =   "*.txt"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   14
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter configuration:"
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "Import black list"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox fSender 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Text            =   "frmAddRule.frx":0442
         Top             =   2760
         Width           =   4695
      End
      Begin VB.TextBox fSubject 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   4695
      End
      Begin VB.TextBox fKeyword 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label6 
         Caption         =   "Seperate addresses with ;"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Sender filter:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Seperate keywords with ;"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Subject line filter:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Seperate keywords with ;"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Keyword parameter(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type of filter:"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.OptionButton Option3 
         Caption         =   "Sender filter"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Subject filter"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Keyword filter"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAddRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim bListBuffer As String
CommonDialog1.ShowOpen
On Error GoTo endSubme
Open CommonDialog1.FileName For Input As #1
Do Until EOF(1)
Line Input #1, bListBuffer
If fSender.Text = "" Then
    fSender.Text = bListBuffer
Else
    fSender.Text = fSender.Text & ";" & bListBuffer
End If
Loop
Close #1
endSubme:
End Sub

Private Sub Command2_Click()
If Option1.value = True Then
   AddKeywordFilter
End If
If Option2.value = True Then
   AddSubjectFilter
End If
If Option3.value = True Then
   AddSenderFilter
End If
DoEvents
Me.Hide
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Option1_Click()
If Option1.value = True Then
   fKeyword.Enabled = True
   fSubject.Enabled = False
   fSender.Enabled = False
   fKeyword.BackColor = &HFFFFFF
   fSubject.BackColor = &HE0E0E0
   fSender.BackColor = &HE0E0E0
   Command1.Enabled = False
End If
End Sub

Private Sub Option2_Click()
If Option2.value = True Then
   fKeyword.Enabled = False
   fSubject.Enabled = True
   fSender.Enabled = False
   fKeyword.BackColor = &HE0E0E0
   fSubject.BackColor = &HFFFFFF
   fSender.BackColor = &HE0E0E0
   Command1.Enabled = False
End If
End Sub

Private Sub Option3_Click()
If Option3.value = True Then
   fKeyword.Enabled = False
   fSubject.Enabled = False
   fSender.Enabled = True
   fKeyword.BackColor = &HE0E0E0
   fSubject.BackColor = &HE0E0E0
   fSender.BackColor = &HFFFFFF
   Command1.Enabled = True
End If
End Sub

Private Sub AddKeywordFilter()
' Add a keyword based filter to
' filter.cfg
Open App.Path & "\Config\filters.cfg" For Append As #1
Print #1, "{"
Print #1, "Rule" & eRuleCount + 1 & ":"
Print #1, "Type: Keyword"
Print #1, "Keywords: " & fKeyword.Text
Print #1, "}"
Close #1
Call frmFilters.LoadRules
End Sub

Private Sub AddSubjectFilter()
' Add a keyword based filter to
' filter.cfg
Open App.Path & "\Config\filters.cfg" For Append As #1
Print #1, "{"
Print #1, "Rule" & eRuleCount + 1 & ":"
Print #1, "Type: Subject"
Print #1, "Subject: " & fSubject.Text
Print #1, "}"
Close #1
Call frmFilters.LoadRules
End Sub

Private Sub AddSenderFilter()
' Add a keyword based filter to
' filter.cfg
Open App.Path & "\Config\filters.cfg" For Append As #1
Print #1, "{"
Print #1, "Rule" & eRuleCount + 1 & ":"
Print #1, "Type: Sender"
Print #1, "Sender: " & fSender.Text
Print #1, "}"
Close #1
Call frmFilters.LoadRules
End Sub
