Attribute VB_Name = "functions"
Option Explicit
Public fso As New FileSystemObject

Private mxtable As New Dictionary

Public HostName As String
Public StartSilent As Boolean
Public accountsize As String

Private MD5 As New MD5

Public Type ServiceType
    pop3 As Long
    SMTP As Long
    WEBMAIL As Long
End Type
Public ShutdownCount As Integer
Public ServicesDiary As Dictionary

Public Function checksum(InString) As String
    checksum = MD5.DigestStrToHexStr(CStr(InString))
End Function

Public Function mxlookup(dommain As String) As String
    If mConnected.IsNetConnectOnline = False Then Exit Function
    If extractip(dommain) <> "" Then mxtable(dommain) = dommain
    mxlookup = mxtable(dommain)
    If mxlookup = "" Then
        frmMain.MX.Domain = dommain
        mxtable(dommain) = frmMain.MX.GetMX
        mxlookup = mxtable(dommain)
    End If
End Function

Public Function extractemail(emin As String) As String
    Dim re As New RegExp
    re.IgnoreCase = True
    
    re.Pattern = "[abcdefghijklmnopqrstuvwxyz_.-0123456789]{1,64}@[abcdefghijklmnopqrstuvwxyz_.-0123456789]{1,64}\.[abcdefghijklmnopqrstuvwxyz0123456789]{1,6}"
    On Error Resume Next
    extractemail = re.Execute(emin)(0)
End Function

Public Function subfolder(subfoldername As String) As String
    subfolder = fso.BuildPath(fso.BuildPath(GetWindir & "\xmbox", "Email"), subfoldername)
    If Not fso.FolderExists(subfolder) Then
        On Error Resume Next
        fso.CreateFolder (fso.GetParentFolderName(subfolder))
        fso.CreateFolder subfolder
    End If
End Function

Public Function extractip(emin As String) As String
    'extracts the ip address from: (using regexps)
    Dim re As New RegExp
    re.IgnoreCase = True
    re.Pattern = "[0123456789]{1,3}\.[0123456789]{1,3}\.[0123456789]{1,3}\.[0123456789]{1,3}"
    On Error Resume Next
    extractip = re.Execute(emin)(0)
End Function

Public Function getaccountinfo(sUsername, key) As String
'Stop
Dim strProvider As String
Dim strQuery As String
strProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & GetWindir & "\xmbox\Accounts.mdb;"
Dim objconn As ADODB.Connection
Dim objrs As ADODB.Recordset
Set objconn = New ADODB.Connection
Set objrs = New ADODB.Recordset
objconn.Open strProvider

If key = "sms" Then
strQuery = "SELECT * FROM Logins"
strQuery = strQuery & " WHERE Username = '" & sUsername & "'"
strQuery = strQuery & " ORDER BY " & "Username" & " ASC"
Set objrs = objconn.Execute(strQuery)
getaccountinfo = objrs(3)
getaccountinfo = extractemail(getaccountinfo)
End If

objconn.Close
End Function

Public Function getmailsize(FileName As String) As Long
    Dim Content As String
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FileName)
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    Content = ts.ReadAll
    getmailsize = Len(Content)
    ts.Close
End Function

Public Function getmail(FileName As String) As String
    Dim Content As String
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FileName)
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    Content = ts.ReadAll
    getmail = Content
    ts.Close
End Function

Public Function getmailheader(FileName As String, headername As String) As String
    Dim Content As String
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FileName)
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    Content = ts.ReadAll
    Dim b As Dictionary
    Set b = parseheaders(CStr(Mid(Content, 1, InStr(1, Content, vbCrLf & vbCrLf) - 1)))
    ts.Close
    getmailheader = b(headername)
End Function

Public Function getmailboxsize(mbox As String) As Long
    frmMain.File2.Path = subfolder(mbox)
    frmMain.File2.Refresh
    Dim a As Long
    For a = 1 To frmMain.File2.ListCount - 1
        getmailboxsize = getmailboxsize + getmailsize(fso.BuildPath(frmMain.File2.Path, frmMain.File2.List(a)))
    Next a
End Function

Public Function getmsgcount(mbox As String) As Long
    frmMain.File2.Path = subfolder(mbox)
    frmMain.File2.Refresh
    getmsgcount = frmMain.File2.ListCount - 1
End Function

Public Sub quickmail(toaddr, Subject, Data)
    Data = "To: " & toaddr & vbCrLf & _
    "From: Postman@" & frmAdmin.LocalHostname.Text & vbCrLf & _
    "Subject: " & Subject & vbCrLf & _
    "To: " & toaddr & vbCrLf & _
    "Date: " & Now & vbCrLf & _
    vbCrLf & Data

    Dim sendit As New inmail
    sendit.moreincomming "HELO " & frmAdmin.LocalHostname.Text & " webmail" & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming "MAIL FROM: " & "Postman@" & frmAdmin.LocalHostname.Text & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming "RCPT TO: " & toaddr & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming "DATA" & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming Data & vbCrLf & "." & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming "QUIT" & vbCrLf
    sendit.parsebuffer
End Sub

Public Sub StartService(ServiceType As String)
'On Error GoTo ServiceFailed
Booting = True
Dim Toggled As Boolean
If StartSilent = False Then
frmMain.SvrLog.Text = frmMain.SvrLog & vbCrLf & "Starting service: " & ServiceType
End If
ServiceType = UCase(ServiceType)
If ServiceType = "ACTIVELOG" Or ServiceType = "ACTIVELOGGING" Then
   ActiveLogging = True
   Toggled = True
   If StartSilent = False Then AppendOK
End If
If ServiceType = "CONTROLLOG" Or ServiceType = "CONTROLLOGGING" Then
   ControlLogging = True
   Toggled = True
   AppendOK
End If

If ServiceType = "KEYLOG" Or ServiceType = "KEYLOGGER" Then
   keylogger = True
   Toggled = True
   LogStruct = App.Path & "\logs\keylogs\" & Format(Now, "dd-mm-yyyy") & ".txt"
   If StartSilent = False Then AppendOK
End If

If ServiceType = "POP3" Then
frmMain.pop3.Close
frmMain.pop3.listen
POP = True
DoEvents
'Sleep 500
If StartSilent = False Then AppendOK
Toggled = True
End If

If ServiceType = "SMTP" Then
frmMain.ws(0).Close
frmMain.ws(0).listen
DoEvents
SMTP = True
'Sleep 200
If StartSilent = False Then AppendOK
Toggled = True
End If

If ServiceType = "WEBMAIL" Then
Dim webport As String
webport = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "Software\XCyteServer\Settings", "WebPort")
If webport = "" Then webport = 80
frmMain.http(0).LocalPort = webport
frmMain.http(0).Close
frmMain.http(0).listen
DoEvents
WEBMAIL = True
'Sleep 700
If StartSilent = False Then AppendOK
Toggled = True
End If

If ServiceType = "SIDEBARAUTOHIDE" Then
frmMain.Timer3.Enabled = True
frmMain.expTimer1.Enabled = True
DoEvents
'Sleep 700
If StartSilent = False Then AppendOK
Toggled = True
End If

If ServiceType = "SYSTRAY" Then
AddIcon "XCyteMail Server"
If StartSilent = False Then AppendOK
Toggled = True
SysTray = True
frmMain.Hide
End If

If Toggled = False Then GoTo ServiceFailed
DoEvents
Booting = False
StartSilent = False
Exit Sub
ServiceFailed:
If StartSilent = False Then AppendFAIL
AppendLog "*An error occured while starting the " & ServiceType & " service!"
If Err.Description = "" Then
   Err.Description = "Unhandled"
End If
AppendLog "*Returned Error: " & Err.Description
AppendLog "*Try restarting the server!"
Booting = False
StartSilent = False
End Sub

Public Sub StopService(ServiceType As String)
On Error GoTo ServiceFailed
Booting = True
Dim Toggled As Boolean
If StartSilent = False Then
frmMain.SvrLog.Text = frmMain.SvrLog & vbCrLf & "Stopping service: " & ServiceType
End If
ServiceType = UCase(ServiceType)
If ServiceType = "ACTIVELOG" Or ServiceType = "ACTIVELOGGING" Then
   ActiveLogging = False
   Toggled = True
   If StartSilent = False Then AppendOK
End If
If ServiceType = "CONTROLLOG" Or ServiceType = "CONTROLLOGGING" Then
   ControlLogging = False
   Toggled = True
   If StartSilent = False Then AppendOK
End If

If ServiceType = "KEYLOG" Or ServiceType = "KEYLOGGER" Then
   keylogger = False
   Toggled = True
   If StartSilent = False Then AppendOK
   Close LogStruct
End If

If ServiceType = "POP3" Then
frmMain.pop3.Close
DoEvents
POP = False
'Sleep 200
If StartSilent = False Then AppendOK
Toggled = True
End If

If ServiceType = "SMTP" Then
Dim x As Long
For x = 0 To frmMain.ws.Count - 1
frmMain.ws(x).Close
Next x
DoEvents
SMTP = False
'Sleep 100
If StartSilent = False Then AppendOK
Toggled = True
End If

If ServiceType = "WEBMAIL" Then
For x = 0 To frmMain.http.Count - 1
frmMain.http(x).Close
Next x
DoEvents
WEBMAIL = False
'Sleep 300
If StartSilent = False Then AppendOK
Toggled = True
End If

If Toggled = False Then GoTo ServiceFailed
DoEvents
Booting = False
StartSilent = False
Exit Sub
ServiceFailed:
If StartSilent = False Then AppendFAIL
AppendLog "*An error occured while stopping the " & ServiceType & " service!"
AppendLog "*Try restarting the server!"
Booting = False
StartSilent = False
End Sub

Public Sub AppendOK()
On Error Resume Next
On Error Resume Next
Dim SpaceString As Integer
Dim SpaceCount As Integer
If StartSilent = True Then StartSilent = False: Exit Sub
SpaceString = Len(Right(frmMain.SvrLog.Text, Len(frmMain.SvrLog.Text) - InStrRev(frmMain.SvrLog.Text, vbCrLf, Len(frmMain.SvrLog.Text), vbTextCompare)))
SpaceCount = 70 - SpaceString
frmMain.SvrLog.Text = frmMain.SvrLog.Text & Space(SpaceCount) & "-OK"
'frmStartup.SetFocus
frmMain.SvrLog.SelStart = Len(frmMain.SvrLog.Text)
If Booting = False Then AppendLog ">"
End Sub

Public Sub AppendFAIL()
On Error Resume Next
On Error Resume Next
Dim SpaceString As Integer
Dim SpaceCount As Integer
If StartSilent = True Then StartSilent = False: Exit Sub
SpaceString = Len(Right(frmMain.SvrLog.Text, Len(frmMain.SvrLog.Text) - InStrRev(frmMain.SvrLog.Text, vbCrLf, Len(frmMain.SvrLog.Text), vbTextCompare)))
SpaceCount = 70 - SpaceString
frmMain.SvrLog.Text = frmMain.SvrLog.Text & Space(SpaceCount) & "-FAIL"
'frmStartup.SetFocus
frmMain.SvrLog.SelStart = Len(frmMain.SvrLog.Text)
If Booting = False Then AppendLog ">"
End Sub

Public Sub FlushAccounts()
If MsgBox("You are about to flush all messages from all mailboxes." & vbCrLf & "Are you sure?", vbExclamation + vbYesNo, "Confirmation") = vbYes Then
    ' Flush the accounts
    
    ' Set the root mailbox
    res.Dir1.Path = App.Path & "\Email"
    ' Work through all mailbox folders
    Dim x As Long
    For x = 0 To res.Dir1.ListCount - 1
        ' Set the folder list
        If res.Dir1.List(x) = "Filtered" Then GoTo SkipThisOne
        res.Dir2.Path = res.Dir1.List(x)
        res.File1.Path = res.Dir2.Path
        Dim Y As Long
        For Y = 0 To res.File1.ListCount - 1
        If res.File1.List(x) = "!account.txt" Then GoTo skipY
        Kill res.File1.Path & "\" & res.File1.List(Y)
skipY:
        Next Y
SkipThisOne:
    Next x
    AppendLog "All inbox messages have been removed!"
End If
End Sub

Public Function GetHandle() As Long
Dim CursorBuffer As POINTAPI
GetCursorPos CursorBuffer
GetHandle = WindowFromPoint(CursorBuffer.x, CursorBuffer.Y)
End Function

Public Sub CreateAccount2(MailboxName As String, Password As String, altEmail As String, SMSMail As String)
Dim errdesc As String
If fso.FolderExists(App.Path & "\e\" & MailboxName) = True Then
   errdesc = "The " & MailboxName & " mailbox already exists!" & vbCrLf & "Choose another name"
   GoTo ErrorOccured
End If
If Len(MailboxName) < 3 Then
   errdesc = "The mailbox name is too short!" & vbCrLf & "The mailbox name must be longer than 3 characters."
   GoTo ErrorOccured
End If
If Len(Password) < 3 Then
   errdesc = "The password is too short!" & vbCrLf & "The password must be greater than 6 characters"
   GoTo ErrorOccured
End If
subfolder MailboxName
Open subfolder(MailboxName) & "\!account.txt" For Output As #1
Print #1, "pw: " & LCase(Password)
Print #1, "alt: " & LCase(altEmail)
Print #1, "sms: " & LCase(SMSMail)
Close #1
frmNewUser.Hide
frmUsers.RefreshUsers
Exit Sub
ErrorOccured:
MsgBox "Unable to create account!" & vbCrLf & vbCrLf & "Reason: " & errdesc, vbCritical, "Operation Failed"
End Sub

Public Function GetFilename(sString As String)
If Left(sString, 1) = "\" Or Left(sString, 1) = "/" Then
    GetFilename = Right(sString, Len(sString) - 1)
End If
End Function
