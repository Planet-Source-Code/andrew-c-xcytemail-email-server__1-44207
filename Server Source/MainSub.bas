Attribute VB_Name = "MainSub"
Public Booting As Boolean
Public Protection As Boolean
Public ServerOnline As Boolean
Option Explicit

Sub Main()
' The following lines are commented to disable the protection system cos it ain't finished yet
'InitProtection
'Do Until ProtectionPassed = True
'DoEvents
'Loop
'Protection = true

'on Error GoTo ErrorOccured

    Dim DBasePath As String
    DBasePath = GetWindir & "\xmbox"
    If fso.FolderExists(DBasePath) = False Then
       MkDir DBasePath
    End If

If fso.FileExists(App.Path & "\XCyte.Bin") = False Then
    frmInit.Show
    frmInit.Timer1.Enabled = True
    Do Until frmInit.SetupOK = True
    DoEvents
    Loop
End If

If fso.FolderExists(App.Path & "\Logs") = False Then
   fso.CreateFolder (App.Path & "\Logs")
End If

If fso.FolderExists(App.Path & "\Config") = False Then
   fso.CreateFolder (App.Path & "\Config")
End If

' Check command line parameters
RunTimeCom.GetParams Command$
ProcessComLines

LoadSettings
frmMain.Show
frmMain.File1.Path = subfolder("out")
DoEvents

'frmMain.Popin
frmMain.SvrLog.Width = frmMain.Width - (frmMain.SideBar.Width + frmMain.SideBar.Left)
frmMain.SvrLog.Text = "XCyteMail network server"
frmMain.SvrLog.Text = frmMain.SvrLog & vbCrLf & "Version: " & App.Major & "." & App.Minor & "." & App.Revision
frmMain.SvrLog.Text = frmMain.SvrLog & vbCrLf & "Developer: Andrew Cranston"
frmMain.SvrLog.Text = frmMain.SvrLog & vbCrLf & "Email: Crano@Hotmail.com"
frmMain.SvrLog.Text = frmMain.SvrLog & vbCrLf & ""

frmMain.SvrLog.Text = frmMain.SvrLog.Text & vbCrLf & "Starting server..."
If Protection = False Then
    frmMain.SvrLog.Text = frmMain.SvrLog.Text & vbCrLf & "Protection scheme is disabled."
Else
    frmMain.SvrLog.Text = frmMain.SvrLog.Text & vbCrLf & "Protection scheme is enabled."
End If

' Run autoexec file
'frmMain.SvrLog.Text = frmMain.SvrLog & vbCrLf & "Executing startup script:"
RunAutoexec
frmMain.SvrLog.Text = frmMain.SvrLog.Text & vbCrLf
DoEvents
Sleep 1000
DoEvents


On Error GoTo NoDNS
mDNS.GetDNSInfo
AppendLog "Using DNS Server(s):"
Dim x As Long
For x = 0 To mDNS.mi_DNSCount
AppendLog "   " & x & ". " & mDNS.sDNS(x)
frmDiag.DNSList.AddItem mDNS.sDNS(x)
Next x
frmDiag.DNSList.Text = mDNS.sDNS(0)
GoTo ContOK
NoDNS:
AppendLog "***No DNS server available"


ContOK:
On Error Resume Next
    
    
AppendLog vbCrLf & "Probing internet connection..."
Dim ProbedOK As Boolean
If mConnected.IsNetConnectViaLAN = True Then
   AppendLog "LAN Connection detected"
   ProbedOK = True
End If
If mConnected.IsNetConnectViaModem = True Then
   AppendLog "Modem Connection detected"
   AppendLog "Warning: Running a mail server on a dialup connection is not recommended"
   ProbedOK = True
End If
If mConnected.IsNetConnectViaProxy = True Then
   AppendLog "Proxy server detected"
   AppendLog "Warning: Running a mail server from behind a proxy server is not recommended"
   AppendLog "         Be sure to configure your proxy server to allow incoming/outgoing"
   AppendLog "         connections to the mailserver on ports: 25, 110 and " & frmAdmin.httpport.Text
   ProbedOK = True
End If

If ProbedOK = False Then
   AppendLog "Internet connection not detected"
   AppendLog "Warning: Outgoing mail will be denied"
   ServerOnline = False
Else
   ServerOnline = True
End If

AppendLog ""
StartServer
Exit Sub

ErrorOccured:
AppendLog "   ***An unhandled error occured in Main()"
AppendLog "        Server will now shutdown..."
DoEvents
Sleep 500
ShutdownServer
End Sub

Public Sub AppendLog(sString As String)
' Append some text to the server log window
If ActiveLogging = False Then Exit Sub
frmMain.SvrLog.Text = frmMain.SvrLog.Text & vbCrLf & sString
frmMain.SvrLog.SelStart = Len(frmMain.SvrLog.Text)
End Sub

Public Sub LoadSettings()
Dim POP3State As String
Dim SMTPState As String
Dim WebMailState As String
Dim sHostname As String
Dim webport As String
webport = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "HTTPPort")
If webport = "" Then webport = "80"
If IsNumeric(webport) = False Then webport = "80"
frmAdmin.httpport.Text = webport
sHostname = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "Hostname")
If sHostname = "" Then sHostname = frmMain.ws(0).LocalIP
HostName = sHostname
accountsize = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XCyteServer\Settings", "Accountsize")
If accountsize = "" Then accountsize = 10242880
If accountsize = "0" Then accountsize = 10242880
POP3State = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "POP3")
SMTPState = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "SMTP")
WebMailState = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "WEBMAIL")
If POP3State = "" Then POP3State = "1"
If SMTPState = "" Then SMTPState = "1"
If WebMailState = "" Then WebMailState = "1"
POP3Service = POP3State
SMTPService = SMTPState
WEBMAILService = WebMailState
'frmMain.POP3Check.value = POP3Service
'frmMain.SMTPCheck.value = SMTPService
'frmMain.aspcheck.value = WEBMAILService
frmAdmin.uMinLength.Text = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "minUsername")
frmAdmin.pMinLength.Text = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "minPassword")
frmAdmin.CookieTimeout.Text = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "CookieExpire")
If frmAdmin.uMinLength.Text = "" Then frmAdmin.uMinLength.Text = "5"
If frmAdmin.pMinLength.Text = "" Then frmAdmin.pMinLength.Text = "3"
If frmAdmin.CookieTimeout.Text = "" Then frmAdmin.CookieTimeout.Text = "30"
frmAdmin.LocalHostName.Text = HostName
frmAdmin.accMin.Text = Trim(accountsize)
Dim DNSTemp As String
Dim dns1 As String
Dim dns2 As String
Dim dns3 As String
Dim dns4 As String
Dim AutoDNSValue As String
AutoDNSValue = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XCyteServer\Settings", "AutoDNS")
If AutoDNSValue = "" Then AutoDNSValue = 1
DNSTemp = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XCyteServer\Settings", "DNSServer")
If DNSTemp = "" Or DNSTemp = " " Then
   frmAdmin.AutoDNS.value = 1
   GoTo SkipDNS
End If
dns1 = Left(DNSTemp, InStr(1, DNSTemp, ".", vbTextCompare) - 1)
DNSTemp = Right(DNSTemp, Len(DNSTemp) - Len(dns1) - 1)
dns2 = Left(DNSTemp, InStr(1, DNSTemp, ".", vbTextCompare) - 1)
DNSTemp = Right(DNSTemp, Len(DNSTemp) - Len(dns2) - 1)
dns3 = Left(DNSTemp, InStr(1, DNSTemp, ".", vbTextCompare) - 1)
DNSTemp = Right(DNSTemp, Len(DNSTemp) - Len(dns3) - 1)
frmAdmin.DNSSet1 = dns1
frmAdmin.DNSSet2 = dns2
frmAdmin.DNSSet3 = dns3
frmAdmin.DNSSet4 = DNSTemp
frmAdmin.UserDNS = dns1 & "." & dns2 & "." & dns3 & "." & DNSTemp
frmAdmin.AutoDNS.value = AutoDNSValue
SkipDNS:
Dim wwwroot As String
wwwroot = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XCyteServer\Settings", "WWWRoot")
If wwwroot = "" Or wwwroot = " " Then
    If fso.FolderExists(App.Path & "\wwwroot") = False Then
        MkDir App.Path & "\wwwroot"
    End If
    wwwroot = App.Path & "\wwwroot"
End If
frmAdmin.Rootpath.Text = wwwroot

Dim mforward As String
mforward = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XCyteServer\Settings", "MailForward")
If mforward = "" Or mforward = " " Then
    mforward = 0
End If
frmAdmin.Mailforward.value = mforward
End Sub

Public Sub SaveSettings()
'RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "POP3", frmMain.POP3Check.value
'RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "SMTP", frmMain.SMTPCheck.value
'RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "WEBMAIL", frmMain.aspcheck.value
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "Hostname", frmAdmin.LocalHostName.Text
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "Accountsize", Str(accountsize)
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "minUsername", frmAdmin.uMinLength.Text
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "minPassword", frmAdmin.pMinLength.Text
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "CookieExpire", frmAdmin.CookieTimeout.Text

RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "Maxbuffer", frmAdmin.MaxBuf.Text
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "DNSServer", frmAdmin.UserDNS
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "AutoDNS", frmAdmin.AutoDNS.value
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "HTTPPort", frmAdmin.httpport.Text

RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "WWWRoot", frmAdmin.Rootpath.Text
RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "MailForward", frmAdmin.Mailforward.value
End Sub

Public Sub ShutdownServer()
Booting = True
AppendLog ""
DoEvents
Sleep 500
AppendLog "Terminating server services..."
DoEvents
Sleep 100
AppendLog "Disconnecting active users..."
DoEvents
Sleep 300
AppendLog "Saving preferences..."
SaveSettings
DoEvents
Sleep 100
DoEvents
StopService "POP3"
Sleep 500
DoEvents
StopService "SMTP"
Sleep 500
DoEvents
StopService "WEBMAIL"
Sleep 500
DoEvents
AppendLog "Disconnecting active users..."
Dim x As Long
For x = 1 To frmMain.ws.Count - 1
frmMain.ws(x).Close
Next x

For x = 1 To frmMain.http.Count - 1
frmMain.http(x).Close
Next x

AppendLog "Shutting down server..."
DoEvents
RemoveIcon
Sleep 1000
End
End Sub

Public Function FileExists(sFilename As String) As Boolean
On Error GoTo NoFile
Open sFilename For Input As #1
Close #1
FileExists = True
Exit Function
NoFile:
FileExists = False
End Function

Public Sub StartServer()
' Start the mail server with all services
Booting = True
DoEvents
'Sleep 500
'If frmMain.POP3Check.value = 1 Then StartService "POP3"
'If frmMain.SMTPCheck.value = 1 Then StartService "SMTP"
'If frmMain.aspcheck.value = 1 Then StartService "WEBMAIL"
'Sleep 800
frmMain.File1.Path = subfolder("out")
Open subfolder("Filtered") & "\!account.txt" For Output As #1
Print #1, "pw: Filtered"
Print #1, "alt: postman@localhost"
Print #1, "sms: postman@localhost"
Close #1
'AppendLog "Active services:"
'Appendservices
'AppendLog vbCrLf
AppendLog "Server successfully started at " & (Format(Now, "ddd dd-mm-yyyy hh:mm:ss AM/PM"))
SysVars("StartTime") = (Format(Now, "ddd dd-mm-yyyy hh:mm:ss AM/PM"))
'AppendLog vbCrLf
AppendLog ">"
Booting = False
ServerActive = True
End Sub

Public Sub StopServer()
' Stop the mail server
frmMain.ws(0).Close
frmMain.http(0).Close
frmMain.pop3.Close
AppendLog "Mail server offline."
AppendLog " "
AppendLog ">"
ServerActive = False
End Sub


