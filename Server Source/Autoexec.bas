Attribute VB_Name = "Autoexec"
' Autoexec file executor
Option Explicit
Dim sString As String
Public SysVars As New Dictionary

Public Sub RunAutoexec()
If FileExists(App.Path & "\Config\Autoexec.txt") = False Then Exit Sub
Open App.Path & "\Config\autoexec.txt" For Input As #1
Do Until EOF(1)
Input #1, sString
ProcessLine sString
Loop
Close #1
End Sub

Private Sub ProcessLine(InputString As String)
If Left(InputString, 1) = ";" Then
   ' Commented line, so skip over
   Exit Sub
End If
If LCase(Left((InputString), Len("StartService"))) = "startservice" Then
   StartService Right(InputString, Len(InputString) - Len("StartService") - 1)
End If

If LCase(Left((InputString), Len("sidebar"))) = "sidebar" Then
   If Right(InputString, Len(InputString) - Len("sidebar") - 1) = "false" Then
      frmMain.SideBar.Left = 0
      
      Autopop = True
   Else
      frmMain.PopOut
      Autopop = False
   End If
End If

If LCase(Left((InputString), Len("message"))) = "message" Then
   frmMain.SvrLog.Text = frmMain.SvrLog.Text & vbCrLf & vbCrLf & Right(InputString, Len(InputString) - Len("message") - 1)
End If

If LCase(Left((InputString), Len("set"))) = "set" Then
    Dim d As String
    Dim t As String
    Dim f As String
    d = Right(InputString, Len(InputString) - Len("set") - 1)
    t = Mid(d, 1, InStr(1, d, "=") - 1)
    f = Mid(d, InStr(1, d, "=") + 1)
    SysVars.Add t, f
End If




End Sub
