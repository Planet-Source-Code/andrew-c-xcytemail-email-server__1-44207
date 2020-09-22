Attribute VB_Name = "GeneralDeclare"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public POP3Service As Long
Public SMTPService As Long
Public WEBMAILService As Long
Public AdminPassword As String
Public Called As Boolean
Public ServerActive As Boolean
Public Autopop As Boolean
