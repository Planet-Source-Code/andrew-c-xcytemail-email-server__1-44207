Attribute VB_Name = "Protection"
'// Protection module
'// Give 30 days to try the software
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public ProtectionPassed As Boolean
Option Explicit

Public Sub InitProtection()
Dim WinDir As String
If fso.FolderExists("C:\Windows") = True Then WinDir = "C:\Windows"
If fso.FolderExists("C:\WinNT") = True Then WinDir = "C:\WinNT"
If FileExists(WinDir & "\media64ti.dll") = False Then
   ' Protection hasn't been initialized
   ' Create recall and deg
   Open WinDir & "\media64ti.dll" For Append As #1
   Print #1, "$(WinSysPathSysFile),,,3/8/99 12:00:00 AM,147728,2.40.4275.1"
   Close #1
   Open WinDir & "\netchronti.dll" For Append As #1
   Print #1, Format(Now, "dd/mm/yyyy")
   Close #1
   ProtectionPassed = False
End If
If FileExists(WinDir & "\media64ti.dll") = True Then
   If FileExists(WinDir & "\netchronti.dll") = False Then
Expired:
      MsgBox "You have exceeded your 30 day trial period!" & vbCrLf & "The software will now stop functioning.", vbCritical, "Protection error"
      End
   Else
      Dim StoredDate As String
      Open WinDir & "\netchronti.dll" For Input As #1
      Input #1, StoredDate
      Close #1
      If DateDiff("d", StoredDate, Format(Now, "dd/mm/yyyy")) > 30 Then GoTo Expired
      If DateDiff("d", StoredDate, Format(Now, "dd/mm/yyyy")) < 0 Then GoTo Expired
   End If
ProtectionPassed = True
End If
End Sub

Public Function GetWindir() As String
Dim PathData As String
PathData = Space(20)
Dim h As Long
h = GetWindowsDirectory(PathData, Len(PathData))
PathData = Trim(PathData)
GetWindir = PathData
' Run a quick check for invalid characters at the end
Const CheckString = "abcdefghijklmnopqrstuvwxyz1234567890.~-_=+"
Dim x As Long
For x = 1 To Len(CheckString)
If LCase(Right(GetWindir, 1)) = Mid(CheckString, x, 1) Then GoTo checkok
Next x
GetWindir = Left(GetWindir, Len(GetWindir) - 1)
checkok:
End Function

