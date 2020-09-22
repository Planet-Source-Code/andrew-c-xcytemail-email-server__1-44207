Attribute VB_Name = "Logging"
Const DatePOS As Integer = 40

Public Sub TelnetLog(TargetIP As String)
If FileExists(App.Path & "\Logs\Telnet.log") = False Then
   ' Setup the file structures
   Open App.Path & "\Logs\Telnet.log" For Output As #1
   Print #1, "XCyteMail server telnet log"
   Print #1, "------------------------------------------------------------------------------------"
   Print #1, "IP Address:               |             Date/Time:"
   Print #1, "------------------------------------------------------------------------------------"
   Close #1
End If

' Make sure we only got an ip address on the TargetIP
If InStr(1, TargetIP, " ", vbTextCompare) <> 0 Then
   TargetIP = Left(TargetIP, InStr(1, TargetIP, " ", vbTextCompare) - 1)
End If

Open App.Path & "\Logs\Telnet.log" For Append As #1
Print #1, TargetIP & Space(DatePOS - Len(TargetIP)) & Format(Now, "dd/mm/yyyy hh:mm:ss AM/PM")
Close #1
End Sub
