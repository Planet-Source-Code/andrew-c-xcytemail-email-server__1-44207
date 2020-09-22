VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class handles 1 incomming smtp session per instance.
'either from another smtp server, a client ie outlook express, or a telnet session.

'it is also used internally (in some occasions), for instance, in the SMS notification
'section, this class will fake an smtp session with another instance of itself, just to simplify things.

Public inbuffer As String
Public outbuffer As String
Public ErrorString2 As String

Public State As smtpinstate

Public fromaddr As String
Public toaddr As New Collection
Public body As String
Public helo As String
Public ClientIP As String
Public ToggleClose As Boolean

Public Sub moreincomming(incoming As String)
    On Error GoTo 0
    inbuffer = inbuffer & incoming
    parsebuffer
    parsebuffer
End Sub

Public Sub parsebuffer()
'Stop
'
'Stop
    If InStr(1, inbuffer, vbCrLf) = 0 Then Exit Sub
    If ControlLogging = True Then AppendLog "Incoming> " & previewtopline
    If previewtopline = "" Then pulltopline: parsebuffer: Exit Sub
    If State <> eDone And LCase(Left(previewtopline, 4)) = "quit" Then
        outbuffer = outbuffer & "250 Bye Bye!"
        ToggleClose = True
        Exit Sub
    End If
    If State <> dGettingdata And LCase(Left(previewtopline, 4)) = "rset" Then
            outbuffer = "250 Ok, reset successful. Message canceled!" & vbCrLf
            fromaddr = ""
            For a = 1 To toaddr.Count
                toaddr.Remove 1
            Next a
            body = ""
            inbuffer = ""
            State = aWanthelo
            Exit Sub
    End If
    Select Case State
    Case smtpinstate.aWanthelo
        If LCase(Left(previewtopline, 4)) = "helo" Then
            State = bWantfrom
            ClientIP = Right(previewtopline, Len(previewtopline) - 5)
            outbuffer = outbuffer & "250 Hello! Welcome to " & frmAdmin.LocalHostname.Text & "'s XCyteMail server!" & vbCrLf
'            TelnetLog ClientIP
 '           AppendLog "Telnet access: " & ClientIP
            helo = Mid(previewtopline, 6)
        ElseIf LCase(Left(previewtopline, 4)) = "ehlo" Then
            outbuffer = outbuffer & "502 No ESMTP support please" & vbCrLf
        ElseIf LCase(Left(previewtopline, 4)) = "help" Then
            outbuffer = outbuffer & "250 State your ip (ie HELO 142.152.168.3)" & vbCrLf
        Else
            outbuffer = outbuffer & "503 Please issue HELO command first" & vbCrLf
            AppendErrors "No HELO Identification Issued"
        End If
        pulltopline
    Case smtpinstate.bWantfrom
        If LCase(Left(previewtopline, 4)) = "mail" And extractemail(previewtopline) <> "" Then
            fromaddr = extractemail(previewtopline)
            outbuffer = outbuffer & "250 " & fromaddr & " sender ok." & vbCrLf
            State = cGettingrcpt
        ElseIf LCase(Left(previewtopline, 4)) = "help" Then
            outbuffer = outbuffer & "250 specify email sender, (ie MAIL FROM: me@myemail.com)" & vbCrLf
        Else
            outbuffer = outbuffer & "503 Expected 'MAIL', see help for example" & vbCrLf
            AppendErrors "Bad sender"
        End If
        pulltopline
    Case smtpinstate.cGettingrcpt
'    Stop
        If LCase(Left(previewtopline, 4)) = "rcpt" And extractemail(previewtopline) <> "" Then
            If toaddr.Count > 60 Then
                outbuffer = outbuffer & "452 Too many recipients" & vbCrLf
                AppendErrors "Too many recipients"
            Else
                If InStr(1, extractemail(previewtopline), CStr("@" & frmAdmin.LocalHostname.Text)) <> 0 Or InStr(1, extractemail(previewtopline), CStr("@localhost")) <> 0 Or InStr(1, extractemail(previewtopline), CStr("@127.0.0.1")) <> 0 Then
                    'ok, it's to one of us, check first to see whether they exist (over size, too big, etc. would all come here)
                    cl = Mid(extractemail(previewtopline), 1, InStr(1, extractemail(previewtopline), "@") - 1)
                    If fso.FolderExists(fso.BuildPath(subfolder(""), cl)) Then
                        If getmailboxsize(CStr(cl)) < accountsize Then
                            outbuffer = outbuffer & "250 " & extractemail(previewtopline) & " Local server account OK" & vbCrLf
                            toaddr.Add extractemail(previewtopline)
                        Else
                            outbuffer = outbuffer & "552 " & extractemail(previewtopline) & " account is full. I will attempt to SMS them." & vbCrLf
                            quickmail getaccountinfo(cl, "sms"), "account overflowing", "Yr " & cl & "@" & frmAdmin.LocalHostname.Text & " email account is full. Please delete some msg's. A msg has been bounced."
                            quickmail getaccountinfo(cl, "alt"), "account overflowing", "Your " & cl & "@" & frmAdmin.LocalHostname.Text & " email account is overflowing. Please delete some msg's to free up space, as an email was just rejected that was addressed to you."
                            AppendErrors "Account is full"
                        End If
                    Else
                        AppendErrors "Account doesn't exist"
                        outbuffer = outbuffer & "502 " & extractemail(previewtopline) & " Account doesn't exist," & vbCrLf
                    End If
                Else
                    '// Check for mail forwarding
                    If frmAdmin.Mailforward.value = 0 Then
                        AppendErrors "Mail forwarding not available"
                        outbuffer = outbuffer & "503 Mail forwarding not available on server" & vbCrLf
                    Else
                        toaddr.Add extractemail(previewtopline)
                        outbuffer = outbuffer & "250 " & extractemail(previewtopline) & " recipient OK. Will forward mail." & vbCrLf
                    End If
                End If
            End If
        ElseIf LCase(Left(previewtopline, 4)) = "help" Then
            outbuffer = outbuffer & "250 specify email recipient(s), 1 per command (ie RCPT TO: me@myemail.com). send the 'DATA' command when done." & vbCrLf
        
        ElseIf LCase(Left(previewtopline, 4)) = "data" Then
            If toaddr.Count > 0 Then
                outbuffer = outbuffer & "354 " & toaddr.Count & " recipient" & IIf(toaddr.Count <> 1, "s", "") & " specified, enter mail data and end with . on a new line" & vbCrLf
                State = dGettingdata
            Else
                outbuffer = outbuffer & "503 No recipient address specified (expected RCPT)" & vbCrLf
                AppendLog "No recipient address specified"
            End If
        ElseIf extractemail(previewtopline) = "" Then
            outbuffer = outbuffer & "501 Bad Email Address on the line: " & previewtopline & vbCrLf
            AppendLog "Bad Email Address"
        Else
            outbuffer = outbuffer & "503 Expected 'RCPT' or 'DATA', see help for example" & vbCrLf
            AppendLog "Expected RCPT or DATA"
        End If
        pulltopline
    Case smtpinstate.dGettingdata
        If Right(inbuffer, 5) = vbCrLf & "." & vbCrLf Then
            body = Mid(inbuffer, 1, InStr(1, inbuffer, vbCrLf & "." & vbCrLf) - 1)
            inbuffer = Mid(inbuffer, InStr(1, inbuffer, vbCrLf & "." & vbCrLf) + 3)
            State = eDone
            outbuffer = outbuffer & "250 Ok, body completed" & vbCrLf
        End If
    Case smtpinstate.eDone
        If Left(previewtopline, 4) = "QUIT" Then
        'Stop
            Dim ts As TextStream
            Dim EmailCode As String
            For a = 1 To toaddr.Count
                fn = fso.BuildPath(subfolder("out"), Timer * 100 & Int(Rnd() * 10000) & ".txt")
                
                If InStr(1, toaddr(a), "@" & frmAdmin.LocalHostname.Text) Then
                    'ok, why dont we just save it in our contacts maildrop, instead of our outbox
                    cl = Mid(toaddr(a), 1, InStr(1, toaddr(a), "@") - 1)
                    EmailCode = Timer * 100 & Int(Rnd() * 10000)
                    fn = fso.BuildPath(subfolder(CStr(cl) & "\Main"), EmailCode & ".txt")
                    If getaccountinfo(cl, "sms") <> "" Then
                        'they are using the cell phone notification system,
                        'so, lets notify their cell phone
                        Dim sendit As New inmail
                        sendit.moreincomming "HELO " & frmAdmin.LocalHostname.Text & " automail" & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming "MAIL FROM: " & "Postman@" & frmAdmin.LocalHostname.Text & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming "RCPT TO: " & getaccountinfo(cl, "sms") & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming "DATA" & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming Left(body, 2000) & vbCrLf & "." & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming "QUIT" & vbCrLf
                        sendit.parsebuffer
                    End If
                End If
                
                Set ts = fso.OpenTextFile(fn, ForWriting, True)
                ts.WriteLine fromaddr
                ts.WriteLine toaddr(a)
                ts.WriteLine "0"
                ts.WriteLine Now
                ts.Write body
                ts.Close
                DoEvents
                                
                ' Add the message into the user's database
                If InStr(1, toaddr(a), "@" & frmAdmin.LocalHostname.Text) Then
                    Dim dUsername As String
                    Dim dTimestamp As String
                    Dim dSubject As String
                    Dim dFrom As String
                    Dim dID As String
                    If InStr(1, "@", toaddr(a), vbTextCompare) = 0 Then
                        dUsername = Left(toaddr(a), InStr(1, toaddr(a), "@", vbTextCompare) - 1)
                    Else
                        dUsername = toaddr(a)
                    End If
                    dTimestamp = Format(Now, "hh:mm:ss DD/MM/YYYY")
                    dSubject = Right(body, Len(body) - InStr(1, LCase(body), "subject:") + 1)
                    dSubject = Left(dSubject, InStr(1, dSubject, vbCrLf, vbTextCompare) - 1)
                    dSubject = Right(dSubject, Len(dSubject) - 9)
                    dFrom = fromaddr
                    ' Uncomment the next line to enable email indexing to access databases
                    'AddEmail dUsername, EmailCode, dSubject, dTimestamp, fromaddr
                End If
                
                
                ' Assemble filtering stuff
                On Error Resume Next
                Dim FilterFilename As String
                FilterFilename = fn
                GoTo SkipFilter
                If FilterMessage(FilterFilename) = True Then
                        Dim sendit2 As New inmail
                        sendit2.moreincomming "HELO " & frmAdmin.LocalHostname.Text & " webmail" & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "MAIL FROM: " & "Postman@" & frmAdmin.LocalHostname.Text & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "RCPT TO: " & fromaddr & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "DATA" & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "to: " & toaddr(a) & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "from: Postman@" & frmAdmin.LocalHostname.Text & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "subject: Email rejected!"
                        sendit2.parsebuffer
                        sendit2.moreincomming vbCrLf & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "Your email message to " & toaddr(a) & " was rejected by the system!" & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "The message contained banned keywords, and was filtered." & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming vbCrLf & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "If you believe this to be a fault, please report it to your local administrator!"
                        sendit2.parsebuffer
                        sendit2.moreincomming vbCrLf & "." & vbCrLf
                        sendit2.parsebuffer
                        sendit2.moreincomming "QUIT" & vbCrLf
                        sendit2.parsebuffer
                End If
SkipFilter:
                ' Message has been filtered
            Next a
            outbuffer = outbuffer & "250 Bye Bye" & vbCrLf
            ToggleClose = True
        ElseIf Left(previewtopline, 4) = "RSET" Then
            outbuffer = "250 Ok, message canceled" & vbCrLf
            fromaddr = ""
            For a = 1 To toaddr.Count
                toaddr.Remove 1
            Next a
            body = ""
            inbuffer = ""
            State = aWanthelo
        Else
            outbuffer = outbuffer & "502 expected either QUIT or RSET" & vbCrLf
        End If
        pulltopline
    End Select
End Sub

Private Function previewtopline() As String
    a = InStr(1, inbuffer, vbCrLf)
    If a > 0 Then
        previewtopline = Mid(inbuffer, 1, a - 1)
    End If
End Function

Private Function pulltopline() As String
    a = InStr(1, inbuffer, vbCrLf)
    If a > 0 Then
        pulltop = Mid(inbuffer, 1, a - 1)
        inbuffer = Mid(inbuffer, a + 2)
    End If
End Function

Private Sub AppendErrors(InputError As String)
MsgBox "Test"
If ErrorString2 = "" Then
    ErrorString2 = InputError
End If
End Sub
