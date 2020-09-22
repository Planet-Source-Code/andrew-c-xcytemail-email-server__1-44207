VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFilters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server side filters: (Note: Filters are not operational in this version)"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmFilters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView List1 
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8916
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
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete All"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove Rule"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Rule"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rules() As String
Dim Void As String
Dim eRuleCount As Integer

Private Sub Command1_Click()
frmAddRule.Option1.value = True
frmAddRule.fKeyword.Text = ""
frmAddRule.fSubject.Text = ""
frmAddRule.fSender.Text = ""
frmAddRule.Show
End Sub

Private Sub Command2_Click()
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

Private Sub Command4_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
LoadRules
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
Case Else
RuleString = "Malformed rule format"
End Select
List1.ListItems.Add , , RuleString
End Sub
