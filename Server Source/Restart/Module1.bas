Attribute VB_Name = "Module1"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10

Sub Main()
' Send a message to the XCyteMail server application
' tell it to close by force...
' Then restart it.
If FindWindow(vbNullString, "XCyteMail Network Server") = 0 Then
If Command$ = "" Then
   MsgBox "XCyteMail server is not running!" & vbCrLf & "Cannot restart.", vbCritical, "Failed"
   End
End If
Else
If Command$ = "" Then
    If MsgBox("You are about to restart the XCyteMail Network Server!" & vbCrLf & "Are you sure?", vbCritical + vbOKCancel, "Confirmation") = vbCancel Then End
End If
   Open App.Path & "\~Close.tmp" For Output As #1
   Close #1
   SendMessage FindWindow(vbNullString, "XCyteMail Network Server"), WM_CLOSE, 1, 1
   Do Until FindWindow(vbNullString, "XCyteMail Network Server") = 0
   DoEvents
   Loop
   ' OK, app has terminated
   ' Now run another instance
    Shell App.Path & "\XCyte.exe"
   End
End If
End Sub
