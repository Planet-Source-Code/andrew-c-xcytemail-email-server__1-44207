Attribute VB_Name = "webmailhtml"
Private Const htmlopen As String = "Content-type: text/html" & vbCrLf & vbCrLf & "<HTML><HEAD><TITLE>"
Private Const bodystart As String = "</TITLE></HEAD><BODY bgcolor=""#CCCCCC""><TABLE width=100% height=100% cols=2 rows=1>" & _
"<tr><td valign=top width=1>" & _
"<img src=""img\logo.png"" width=""100"" height=""100""><P><A HREF=""inbox.asp"">Inbox</A><BR><A HREF=""compose.asp"">Compose</A><BR><A HREF=""address.asp"">Address book</A><BR><A HREF=""settings.asp"">Settings</A><BR><A HREF=""logout.asp"">Logout</A><BR><A HREF=""signup.asp"">Signup</A></td><td valign=top>"
Private Const bodyend As String = "</td></tr></table>"
Public FoundPageToggle As Boolean
Public WindowTitle As String
Public CurrentUser As String
Dim Requested As Boolean

Public Function dowebsite(FileName, vars As Dictionary, headers As Dictionary, ByRef pageheader, wsIndex As Integer) As String
'Stop
FileName = LCase(FileName)
WindowTitle = "XCyteMail"
FoundPageToggle = False


'// Check for page redirection
'Stop
If headers("Host") <> frmAdmin.LocalHostname.Text Then
    '// Redirect the user to the real page
    dowebsite = AssemblePage("<P>Redirecting...</P>" & vbNewLine & printjavascript)
    'Exit Function
End If
If FileName = "home.asp" Then
    dowebsite = GetFile(frmAdmin.Rootpath.Text & "\home.asp") ' Load the file directly into the webserver
    FoundPageToggle = True 'Must toggle this otherwise the server will think the page can't be found
    GoTo SkipAll 'Skip all webmail features and send the page
End If


'Stop
If ValidatePassword(vars("un"), vars("pw")) = True And Len(vars("pw")) > 0 Then
        pageheader = pageheader & "Set-Cookie: un=" & vars("un") & "; expires=Fri 28-Jun-2012 13:25:03 GMT;  path=/; domain=" & frmAdmin.LocalHostname.Text & ";" & vbCrLf
        pageheader = pageheader & "Set-Cookie: pw=" & vars("pw") & "; expires=Fri 28-Jun-2012 13:25:03 GMT;  path=/; domain=" & frmAdmin.LocalHostname.Text & ";" & vbCrLf
        ' Is it a requested page that exists?
        If fso.FileExists(frmAdmin.Rootpath.Text & "\" & FileName) = True Then
            Dim FileHandle01 As Long
            Dim TempBuffer As String
            FileHandle01 = FreeFile
            Open frmAdmin.Rootpath.Text & "\" & FileName For Input As #FileHandle01
            Do Until EOF(FileHandle01)
            Input #FileHandle01, TempBuffer
            dowebsite = dowebsite & vbCrLf & TempBuffer
            Loop
            Close FileHandle01
            Requested = True
            FoundPageToggle = True
        End If
End If

If FileName = "login.asp" Or FileName = "authuser.asp" Then
If ValidatePassword(vars("un"), vars("pw")) = True And Len(vars("pw")) > 0 Then
        pageheader = pageheader & "Set-Cookie: un=" & vars("un") & "; expires'=Fri 28-Jun-2012 13:25:03 GMT;  path=/; domain=" & frmAdmin.LocalHostname.Text & ";" & vbCrLf
        pageheader = pageheader & "Set-Cookie: pw=" & vars("pw") & "; expires=Fri 28-Jun-2012 13:25:03 GMT;  path=/; domain=" & frmAdmin.LocalHostname.Text & ";" & vbCrLf
                
        GoTo AssumeInbox
End If
End If

If LCase(FileName) = "authuser.asp" Then
    If Len(vars("pw")) = 0 Then GoTo gotohere2
        If ValidatePassword(vars("un"), vars("pw")) = False And Len(vars("pw") > 0) Then
            Dim PageBuffer As String
            PageBuffer = "        <div align=""center"">"
            PageBuffer = PageBuffer & vbCrLf & "<p>&nbsp;</p>"
            PageBuffer = PageBuffer & vbCrLf & "<table width=""70%"" border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor=""#999999"">"
            PageBuffer = PageBuffer & vbCrLf & "    <tr>"
            PageBuffer = PageBuffer & vbCrLf & "      <td>"
            PageBuffer = PageBuffer & vbCrLf & "        <div align=""center""><font color=""#000000"" face=""Arial, Helvetica, sans-serif""><b><font color=""#FFFFFF"">Login Error:</font></b></font></div>"
            PageBuffer = PageBuffer & vbCrLf & "      </td>"
            PageBuffer = PageBuffer & vbCrLf & "    </tr>"
            PageBuffer = PageBuffer & vbCrLf & "  </table>"
            PageBuffer = PageBuffer & vbCrLf & "  <B>The following error occured:<br>"
            PageBuffer = PageBuffer & vbCrLf & "  <font size=""5"" face=""Arial, Helvetica, sans-serif"" color=""#FF0000"">Username and/or password was not accepted by the server.</font><br>"
            PageBuffer = PageBuffer & vbCrLf & "  <br>"
            PageBuffer = PageBuffer & vbCrLf & "  <a href=""login.asp"">Click here to try again</a><Replace></b></div>"
            dowebsite = AssemblePage(PageBuffer)
            Exit Function
        End If
    If ValidatePassword(vars("un"), vars("pw")) = True And Len(vars("pw")) > 0 Then
            '// Log the user in
            pageheader = pageheader & "Set-Cookie: un=" & vars("un") & "; expires=" & Format(Now + (1 / 24), "ddd dd-mmm-yyyy h:mm:ss") & " GMT;  path=/; domain=" & frmAdmin.LocalHostname.Text & ";" & vbCrLf
            pageheader = pageheader & "Set-Cookie: pw=" & vars("pw") & "; expires=" & Format(Now + (1 / 24), "ddd dd-mmm-yyyy h:mm:ss") & " GMT;  path=/; domain=" & frmAdmin.LocalHostname.Text & ";" & vbCrLf
            FoundPageToggle = True
            GoTo AssumeInbox
    Else
gotohere2:
            If Not (FileName = "login.asp" Or FileName = "signup.asp") Then
dLogin:
                dowebsite = GetFile(frmAdmin.Rootpath & "\login.htm")
                On Error Resume Next
                Exit Function
            End If
    End If
End If

AlreadyLogged:
If Not (FileName = "login.asp" Or FileName = "signup.asp") Then
'// Perform a validation on the user
If vars("un") = "" Or vars("un") = vbEmpty Then
    GoTo dLogin
End If
If ValidatePassword(vars("un"), vars("pw")) = False Then
    GoTo dLogin
End If
End If
    
    
    If FileName = "login.asp" Then
        dowebsite = dowebsite & vbCrLf & showloginpage
        FoundPageToggle = True
'        Exit Function
    End If
    If FileName = "inbox.asp" Then
AssumeInbox:
If vars("folder") = "" Then vars("folder") = "main"
If fso.FolderExists(GetWindir & "\xmbox\email\" & vars("un") & "\Trash") = False Then MkDir (GetWindir & "\xmbox\email\" & vars("un") & "\Trash")
    frmMain.File2.Path = GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("folder")

        If vars("Delete") = "Delete" Then
            'the user wishes to delete some messages before the mail list is shown
            For Each f In vars.Keys
                If vars(f) = "dm" Then
                    On Error Resume Next
                    fso.DeleteFile fso.BuildPath(subfolder(vars("un")), f & ".txt"), True
                    On Error GoTo 0
                End If
            Next
        End If

        accmsgcount = getmsgcount(vars("un"))
        accsize = getmailboxsize(vars("un"))
        
        'Stop
        If fso.FolderExists(GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("folder")) = False Then
            dowebsite = GetFile(frmAdmin.Rootpath & "\foldererror.asp")
            FoundPageToggle = True
            GoTo SkipInbox
        End If
        frmMain.File2.Path = GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("folder")
        For a = 0 To frmMain.File2.ListCount - 1
        Dim eid As String
        Dim fn As String
        Dim sDate As String
        Dim sFrom As String
        Dim sSubject As String

        eid = Left(frmMain.File2.List(a), Len(frmMain.File2.List(a)) - 4)
        fn = fso.BuildPath(frmMain.File2.Path, frmMain.File2.List(a))
        sDate = Left(getmailheader(fn, "Date"), InStr(1, getmailheader(fn, "Date"), " ", vbTextCompare) - 1)
        dowebsite = dowebsite & AttachMailInsert(getmailheader(CStr(fn), "from"), getmailheader(CStr(fn), "Subject"), sDate, getmailsize(CStr(fn)), eid, vars("folder"))
            'dowebsite = dowebsite & "       <tr bgcolor=999999><td><input type=checkbox name=" & eid & " value=dm></td><td>" & getmailheader(CStr(fn), "from") & "</td><td><A href=""getmsg.asp?msg=" & eid & """>" & getmailheader(CStr(fn), "subject") & "</a></td><td>" & getmailsize(CStr(fn)) & "</td></tr>" & vbCrLf
        Next a
        
        FoundPageToggle = True
        WindowTitle = UCase(Left(vars("un"), 1)) & Right(vars("un"), Len(vars("un")) - 1) & "'s mailbox"
        dowebsite = AssembleInbox(dowebsite)
SkipInbox:
    End If
    
    If FileName = "next.asp" Then
    'Stop
    Dim tmpbfr As Long
    Dim NextMSG As String
    frmMain.File2.Path = GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("folder")
    For tmpbfr = 0 To frmMain.File2.ListCount - 1
    If frmMain.File2.List(tmpbfr) = vars("msg") & ".txt" Then
        If tmpbfr = frmMain.File2.ListCount - 1 Then
            vars("msg") = ""
            GoTo AssumeInbox
        End If
        NextMSG = Left(frmMain.File2.List(tmpbfr + 1), Len(frmMain.File2.List(tmpbfr + 1)) - 4)
        vars("msg") = NextMSG
        FoundPageToggle = True
        GoTo GetNext
    End If
    Next tmpbfr
    End If
    
    If FileName = "prev.asp" Then
    'Stop
    frmMain.File2.Path = GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("folder")
    For tmpbfr = 0 To frmMain.File2.ListCount - 1
    If frmMain.File2.List(tmpbfr) = vars("msg") & ".txt" Then
        If tmpbfr = 0 Then
            vars("msg") = ""
            GoTo AssumeInbox
        End If
        NextMSG = Left(frmMain.File2.List(tmpbfr - 1), Len(frmMain.File2.List(tmpbfr - 1)) - 4)
        vars("msg") = NextMSG
        FoundPageToggle = True
        GoTo GetNext
    End If
    Next tmpbfr
    End If
    
    If FileName = "inboxutils.asp" Then
    If vars("Reply") = "Reply" Then
    dowebsite = MakeRedir("compose.asp?To=" & vars("MsgFrom") & "&Subject=Re: " & vars("MsgSubject") & "&PreMSG=" & vars("MessageID") & "&InboxFolder=" & vars("InboxFolder"))
    FoundPageToggle = True
    End If
    
    If vars("Forward") = "Forward" Then
    dowebsite = MakeRedir("compose.asp?To=" & "&Subject=Fwd: " & vars("MsgSubject") & "&PreMSG=" & vars("MessageID") & "&InboxFolder=" & vars("InboxFolder"))
    FoundPageToggle = True
    End If
    
    If vars("Trash") = "Send to trash" Then
        dowebsite = MakeRedir("move.asp?MessageID=" & vars("MessageID") & "&srcfolder=" & vars("InboxFolder") & "&dest=trash")
        FoundPageToggle = True
    End If
    If vars("Trash") = "Empty" Then
        ' Empty the trash folder
        Dim te As New FileSystemObject
        Dim dstfolder As String
        If fso.FolderExists(GetWindir & "\xmbox\email\" & vars("un") & "\Trash") = False Then MkDir (GetWindir & "\xmbox\email\" & vars("un") & "\Trash")
        If fso.FolderExists(GetWindir & "\xmbox\Recycled") = False Then MkDir (GetWindir & "\recycled")
        If fso.FolderExists(GetWindir & "\xmbox\recycled\" & vars("un")) = False Then MkDir (GetWindir & "\xmbox\recycled\" & vars("un"))
        If fso.FolderExists(GetWindir & "\xmbox\recycled\" & vars("un") & "\" & Format(Now, "dd mmm yyyy")) = False Then MkDir (GetWindir & "\xmbox\recycled\" & vars("un") & "\" & Format(Now, "dd mmm yyyy"))
        dstfolder = GetWindir & "\xmbox\recycled\" & vars("un") & "\"
        te.MoveFolder GetWindir & "\xmbox\email\" & vars("un") & "\trash", CStr(dstfolder)
        MkDir GetWindir & "\xmbox\email\" & vars("un") & "trash"
        dowebsite = MakeRedir("inbox.asp?folder=Main")
        FoundPageToggle = True
    End If
    End If
    
    If FileName = "move.asp" Then
'    Stop
        ' Move the message to the specified folder
        Dim pagedata As String
        If fso.FileExists(GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("srcfolder") & "\" & vars("MessageID") & ".txt") = False Then
            pagedata = GetFile(frmAdmin.Rootpath.Text & "\errormessage.asp")
            pagedata = ReplaceVars(pagedata, "$ErrorMessage$", "Message not found on server!")
            pagedata = ReplaceVars(pagedata, "$CallingPage$", "/")
            dowebsite = pagedata
            FoundPageToggle = True
            GoTo skipmove
        End If
    
        Dim mHandle0 As String
        Dim mHandle1 As String
        Dim MailStream As String
        mHandle0 = GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("srcfolder") & "\" & vars("MessageID") & ".txt"
        mHandle1 = GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("dest") & "\" & vars("MessageID") & ".txt"
        
        Dim iStream As TextStream
        Dim oStream As TextStream
        Set iStream = fso.OpenTextFile(mHandle0)
        MailStream = iStream.ReadAll
        iStream.Close
        
        Set oStream = fso.OpenTextFile(mHandle1, ForWriting, True)
        oStream.Write MailStream
        oStream.Close
        
        Kill mHandle0
        DoEvents
        dowebsite = MakeRedir("inbox.asp")
        FoundPageToggle = True
skipmove:
    End If
    
    
    If FileName = "getmsg.asp" Then
GetNext:
    'Stop
        If fso.FileExists(fso.BuildPath(subfolder(vars("un") & "\" & vars("folder")), vars("msg") & ".txt")) = False Then
            pagedata = GetFile(frmAdmin.Rootpath.Text & "\errormessage.asp")
            pagedata = ReplaceVars(pagedata, "$ErrorMessage$", "The requested message was not found on the server!")
            pagedata = ReplaceVars(pagedata, "$CallingPage$", vars("Referer"))
            GoTo skipallsomethingbadhappened
        End If
        
        Source = getmail(fso.BuildPath(subfolder(vars("un") & "\" & vars("folder")), vars("msg") & ".txt"))
        mh = Left(Source, InStr(1, Source, vbCrLf & vbCrLf) - 1)
        body = Mid(Source, InStr(1, Source, vbCrLf & vbCrLf) + 3)
        Set hlist = parseheaders(CStr(mh))
        pagedata = GetFile(frmAdmin.Rootpath & "\viewmsg.htm")
        pagedata = ReplaceVars(pagedata, "$msgto$", CStr(hlist("to")))
        pagedata = ReplaceVars(pagedata, "$msgfrom$", CStr(hlist("from")))
        pagedata = ReplaceVars(pagedata, "$msgsubject$", CStr(hlist("subject")))
        pagedata = ReplaceVars(pagedata, "$msgcc$", CStr(hlist("cc")))
        pagedata = ReplaceVars(pagedata, "$sentdate$", CStr(hlist("date")))
        pagedata = ReplaceVars(pagedata, "$messagedata$", CStr(body))
        pagedata = ReplaceVars(pagedata, "$messageid$", vars("msg"))
        dowebsite = pagedata
skipallsomethingbadhappened:
        FoundPageToggle = True
    End If
    
    If FileName = "reply.asp" Then
    
    End If
    
    
    If FileName = "software.asp" Then
        dowebsite = "<DIV Align=""Left"">"
        dowebsite = dowebsite & vbCrLf & "<B>Server running:</B> XCyteHybrid 1.1.1"
        dowebsite = dowebsite & vbCrLf & "<P><B>Server version:</B> " & App.Major & "." & App.Minor & "." & App.Revision
        dowebsite = dowebsite & vbCrLf & "<P>"
        dowebsite = dowebsite & vbCrLf & "<B>Active Services:</B></br>"
        dowebsite = dowebsite & vbCrLf & "POP3: " & Services.POP & "</br>"
        dowebsite = dowebsite & vbCrLf & "SMTP: " & Services.SMTP & "</br>"
        dowebsite = dowebsite & vbCrLf & "WEBMAIL: " & Services.WEBMAIL & "</br>"
        dowebsite = dowebsite & vbCrLf & "SYSTRAY: " & Services.SysTray & "</br>"
        dowebsite = dowebsite & vbCrLf & "LogStruct: " & Services.LogStruct & "</br>"
        dowebsite = dowebsite & vbCrLf & "KEYLOG: " & Services.keylogger & "</br>"
        dowebsite = dowebsite & vbCrLf & "Activelog: " & Services.ActiveLogging & "</br>"
        dowebsite = dowebsite & vbCrLf & "Controllog: " & Services.ControlLogging & "</br>"
        dowebsite = AssemblePage(dowebsite)
        FoundPageToggle = True
    End If
    
    If FileName = "logout.asp" Then
        ' Remove the user from the frmhtml list
        dowebsite = GetFile(frmAdmin.Rootpath & "\loggedout.htm")
        pageheader = pageheader & "Set-Cookie: un=" & vars("un") & "; expires=Fri 28-Jun-1900 13:25:03 GMT;  path=/; domain=" & frmAdmin.LocalHostname.Text & ";" & vbCrLf
        pageheader = pageheader & "Set-Cookie: pw=" & vars("pw") & "; expires=Fri 28-Jun-1900 13:25:03 GMT;  path=/; domain=" & frmAdmin.LocalHostname.Text & ";" & vbCrLf
        FoundPageToggle = True
    End If
    If FileName = "compose.asp" Then
        'Stop
        CurrentUser = vars("un")
        dowebsite = GetFile(frmAdmin.Rootpath.Text & "\compose.htm")
        dowebsite = ReplaceVars(dowebsite, "$ToVal$", vars("To"))
        dowebsite = ReplaceVars(dowebsite, "$SubjectVal$", vars("Subject"))
        If vars("Body") <> "" Then
        dowebsite = ReplaceVars(dowebsite, "$BodyVal$", vbCrLf & vbCrLf & vbCrLf & "Original Message:" & vbCrLf & " >" & vars("Body"))
        Else
        dowebsite = ReplaceVars(dowebsite, "$BodyVal$", "")
        End If
        If vars("PreMSG") <> "" Then
        dowebsite = ReplaceVars(dowebsite, "$BodyVal$", vbCrLf & vbCrLf & vbCrLf & "Original Message:" & vbCrLf & " >" & LoadFile(GetWindir & "\xmbox\email\" & vars("un") & "\" & vars("InboxFolder") & "\" & vars("PreMSG") & ".txt"))
        Else
        dowebsite = ReplaceVars(dowebsite, "$BodyVal$", "")
        End If
        
        FoundPageToggle = True
    End If
    
    If FileName = "composenoaddress.asp" Then
        dowebsite = GetFile(frmAdmin.Rootpath.Text & "\compose2.htm")
        FoundPageToggle = True
    End If
    
    If FileName = "address.asp" Then
        dowebsite = "<B>Address book not done yet!</B>"
        dowebsite = AssemblePage(dowebsite)
        FoundPageToggle = True
    End If
    
    If FileName = "send.asp" Then
        from = vars("un") & "@" & frmAdmin.LocalHostname.Text
        ato = vars("to")
        Subject = vars("subject")
        body = vars("body")
        
        Data = "Recieved: " & frmAdmin.LocalHostname.Text & " webmail, user=" & vars("un") & vbCrLf & _
        "From: " & from & vbCrLf & _
        "To: " & ato & vbCrLf & _
        "Date: " & Now & vbCrLf & _
        "Subject: " & Subject & vbCrLf & vbCrLf & body
    
        'creates a new instance of the class that is used to control new mail arivals
        'and creates a new instance of it, pluging in the data from the website as oposed
        'to from a socket
        
        Dim sendit As New inmail
        sendit.moreincomming "HELO " & frmAdmin.LocalHostname.Text & " webmail" & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "MAIL FROM: " & from & vbCrLf
        sendit.parsebuffer
        For Each r In Split(ato, ",")
            sendit.moreincomming "RCPT TO: " & r & vbCrLf
            sendit.parsebuffer
        Next
        sendit.moreincomming "DATA" & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming Data & vbCrLf & "." & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "QUIT" & vbCrLf
        sendit.parsebuffer
        If sendit.ErrorString2 <> "" Then
            ' Some sorta sending error occured
            dowebsite = "<P>An error occured while sending your message:</P>" & vbNewLine & "<B>" & sendit.ErrorString2 & "</B>"
            dowebsite = dowebsite & vbCrLf & "<table width=""60%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
            dowebsite = dowebsite & vbCrLf & "  <tr>"
            dowebsite = dowebsite & vbCrLf & "    <td>"
            dowebsite = dowebsite & vbCrLf & "      <form name=""form1"" method=""post"" action=""inbox.asp"">"
            dowebsite = dowebsite & vbCrLf & "        <div align=""right"">"
            dowebsite = dowebsite & vbCrLf & "          <input type=""submit"" name=""RetMail"" value=""Return"">"
            dowebsite = dowebsite & vbCrLf & "        </div>"
            dowebsite = dowebsite & vbCrLf & "      </form>"
            dowebsite = dowebsite & vbCrLf & "    </td>"
            dowebsite = dowebsite & vbCrLf & "  </tr>"
            dowebsite = dowebsite & vbCrLf & "</table>"
            GoTo gotohere
      End If
        
        dowebsite = "<P>Your message has been sent to the following recipients:</P>" & vbNewLine & "<B>" & ato & "</B>"
        dowebsite = dowebsite & vbCrLf & "<table width=""60%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
        dowebsite = dowebsite & vbCrLf & "  <tr>"
        dowebsite = dowebsite & vbCrLf & "    <td>"
        dowebsite = dowebsite & vbCrLf & "      <form name=""form1"" method=""post"" action=""inbox.asp"">"
        dowebsite = dowebsite & vbCrLf & "        <div align=""right"">"
        dowebsite = dowebsite & vbCrLf & "          <input type=""submit"" name=""RetMail"" value=""Return"">"
        dowebsite = dowebsite & vbCrLf & "        </div>"
        dowebsite = dowebsite & vbCrLf & "      </form>"
        dowebsite = dowebsite & vbCrLf & "    </td>"
        dowebsite = dowebsite & vbCrLf & "  </tr>"
        dowebsite = dowebsite & vbCrLf & "</table>"
        
gotohere:
        dowebsite = AssemblePage(dowebsite)
        FoundPageToggle = True
    End If
    
    If FileName = "signup.asp" Then
        If vars("Signup") = "Signup" Then
            If Len(vars("pw1")) < frmAdmin.pMinLength.Text Then
                er = "Unable to create account." & vbCrLf & "Reason:" & vbCrLf & "Password is too short"
                GoTo nope
            End If
            If Len(vars("unp")) < frmAdmin.uMinLength.Text Then
                er = "Unable to create account." & vbCrLf & "Reason:" & vbCrLf & "Username is too short"
                GoTo nope
            End If
            If AccountExists(vars("unp")) Then
                er = "Unable to create account." & vbCrLf & "Reason:" & vbCrLf & "Username is all ready taken"
                GoTo nope
            End If
            If vars("pw1") <> vars("pw2") Then
                er = "Unable to create account." & vbCrLf & "Reason:" & vbCrLf & "Passwords didn't match"
                GoTo nope
            End If
            
            'everything is fine, create their account
            CreateAccount vars("unp"), vars("pw1"), vars("alt"), vars("sms")
            
            Data = "From: Postman@" & frmAdmin.LocalHostname.Text & vbCrLf & _
            "To: New  User" & vbCrLf & _
            "Subject: Welcome!" & vbCrLf & _
            "Date: " & Now & vbCrLf & vbCrLf & _
            "Welcome to XCyteMail" & vbCrLf & _
            "This is an automated email message to inform you of your account details." & vbCrLf & _
            "Do not respond to this email." & vbCrLf & vbCrLf & _
            "Account Details:" & vbCrLf & _
            "Username: " & vars("unp") & vbCrLf & _
            "Password: " & vars("pw1") & vbCrLf & vbCrLf & _
            "Server Details:" & vbCrLf & _
            "Server IP: " & frmMain.ws(0).LocalIP & vbCrLf & _
            "POP3 server: " & frmAdmin.LocalHostname.Text & vbCrLf & _
            "SMTP server: " & frmAdmin.LocalHostname.Text & vbCrLf & vbCrLf & _
            "WEBMAIL address: http://" & frmAdmin.LocalHostname.Text & vbCrLf & vbCrLf & _
            "Thank you for using XCyteMail." & vbCrLf & _
            "Have a nice day," & vbCrLf & vbCrLf & _
            "                   The XCyteMail team!" & vbCrLf & vbCrLf
            
            'send them an email welcoming them
            Set sendit = New inmail
            sendit.moreincomming "HELO " & frmAdmin.LocalHostname.Text & " webmail" & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "MAIL FROM: " & "Welcome! <Postman@" & frmAdmin.LocalHostname.Text & ">" & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "RCPT TO: " & vars("unp") & "@" & frmAdmin.LocalHostname.Text & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "RCPT TO: " & vars("alt") & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "DATA" & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming Data & vbCrLf & "." & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "QUIT" & vbCrLf
            sendit.parsebuffer
            
            dowebsite = "Signup sucessful, <A HREF=""login.asp"">Login</A> to continue."
            dowebsite = AssemblePage(dowebsite)
        Else
nope:
            dowebsite = "<P>To sign up to XCyteMail, please enter the following details:<P>" & _
            "<FORM action=""signup.asp"" method=post>" & _
            "<FONT SIZE=3 COLOR=FF0000>" & er & "</FONT>" & _
            "<TABLE>" & _
            "<tr><td>Username:</td><td><INPUT name=unp></td></tr>" & _
            "<tr><td>Password:</td><td><INPUT name=pw1 type=password></td></tr>" & _
            "<tr><td>Confirm:</td><td><INPUT name=pw2 type=password></td></tr>" & _
            "<tr><td colspan=2> &nbsp </td></tr>" & _
            "<tr><td>Alternate email:</td><td><INPUT name=alt></td></tr>" & _
            "<tr><td>SMS email:</td><td><INPUT name=sms></td></tr>" & _
            "</TABLE><INPUT type=submit name=Signup value=Signup></FORM>"
            dowebsite = AssemblePage(dowebsite)
        End If
        FoundPageToggle = True
    End If
    
    If FileName = "settings.asp" Then
        dowebsite = "<B>Modify your settings</B>" & vbCrLf
        
        If vars("Save") = "Save" Then
            If vars("mod") = "pw" Then
                If vars("pw1") <> vars("pw") Then
                    dowebsite = dowebsite & "<FONT size=5 color=FF0000>Enter your current password</FONT>"
                    GoTo no
                End If
                If vars("pw2") <> vars("pw3") Then
                    dowebsite = dowebsite & "<FONT size=5 color=FF0000>Those passwords dont match</FONT>"
                    GoTo no
                End If
                If Len(vars("pw2")) < 3 Then
                    dowebsite = dowebsite & "<FONT size=5 color=FF0000>Password must be 3 letters long</FONT>"
                    GoTo no
                End If
                Set ts = fso.OpenTextFile(fso.BuildPath(subfolder(vars("un")), "!account.txt"), ForAppending, True)
                ts.WriteLine "pw: " & vars("pw3")
                dowebsite = dowebsite & "<FONT size=5 color=FF0000>Password changed, please log back in</FONT>"
                ts.Close
            End If
            If vars("mod") = "alt" Then
                Set ts = fso.OpenTextFile(fso.BuildPath(subfolder(vars("un")), "!account.txt"), ForAppending, True)
                ts.WriteLine "alt: " & vars("alt")
                dowebsite = dowebsite & "<FONT size=5 color=FF0000>Alternate email changed</FONT>"
                ts.Close
            End If
            If vars("mod") = "sms" Then
                Set ts = fso.OpenTextFile(fso.BuildPath(subfolder(vars("un")), "!account.txt"), ForAppending, True)
                ts.WriteLine "sms: " & vars("sms")
                dowebsite = dowebsite & "<FONT size=5 color=FF0000>SMS email changed</FONT>"
                ts.Close
            End If
            vars("mod") = ""
        End If
no:
        dowebsite = dowebsite & "<FORM action=""settings.asp"" method=POST><TABLE>"
        If vars("mod") = "pw" Then
            dowebsite = dowebsite & "<tr><td align=right><b>Old password</b>:</td><td><INPUT name=pw1 type=password></td></tr>" & vbCrLf
            dowebsite = dowebsite & "<tr><td align=right><b>New password</b>:</td><td><INPUT name=pw2 type=password></td></tr>" & vbCrLf
            dowebsite = dowebsite & "<tr><td align=right><b>Confirm</b>:</td><td><INPUT name=pw3 type=password></td></tr>" & vbCrLf
        Else
            dowebsite = dowebsite & "<tr><td align=right><b>Password</b>:</td><td>" & String(Len(vars("pw")), "*") & " <small><A HREF=""settings.asp?mod=pw"">Change</A></small></td></tr>" & vbCrLf
        End If
        
        If vars("mod") = "alt" Then
            dowebsite = dowebsite & "<tr><td align=right><b>Alternate email</b>:</td><td><INPUT name=alt value=" & getaccountinfo(vars("un"), "alt") & "></td></tr>" & vbCrLf
        Else
            dowebsite = dowebsite & "<tr><td align=right><b>Alternate email</b>:</td><td>" & getaccountinfo(vars("un"), "alt") & " <small><A HREF=""settings.asp?mod=alt"">Change</A></small></td></tr>" & vbCrLf
        End If
        
        If vars("mod") = "sms" Then
            dowebsite = dowebsite & "<tr><td align=right><b>SMS email</b>:</td><td><INPUT name=sms value=" & getaccountinfo(vars("un"), "sms") & "></td></tr>" & vbCrLf
        Else
            dowebsite = dowebsite & "<tr><td align=right><b>SMS email</b>:</td><td>" & getaccountinfo(vars("un"), "sms") & " <small><A HREF=""settings.asp?mod=sms"">Change</A></small></td></tr>" & vbCrLf
        End If
        dowebsite = dowebsite & "</TABLE>"
        If vars("mod") <> "" Then dowebsite = dowebsite & "<INPUT type=hidden name=mod value=" & vars("mod") & "><INPUT type=submit name=Save value=Save></FORM>"
        dowebsite = AssemblePage(dowebsite)
        End If
'        Stop
SkipAll:
'    FoundPageToggle = True
On Error GoTo errorocc
    If FoundPageToggle = False Then
        ' Page not found. AKA: 404 Error
        dowebsite = GetFile(frmAdmin.Rootpath.Text & "\404.htm")
    End If
    FoundPageToggle = False
    'Stop
    ' $Currentuser$ variable
    If InStr(1, dowebsite, "$Currentuser$", vbTextCompare) <> 0 Then
        If vars("un") = " " Or vars("un") = "" Then
        dowebsite = ReplaceVars(dowebsite, "$Currentuser$", "Not Logged In")
        Else
        dowebsite = ReplaceVars(dowebsite, "$Currentuser$", vars("un"))
        End If
    End If
            
    ' $InboxFolder$ variable
    dowebsite = ReplaceVars(dowebsite, "$InboxFolder$", vars("folder"))
            
            
            
    ' $DateTime$ variable
    dowebsite = ReplaceVars(dowebsite, "$DateTime$", Format(Now, "dd/mm/yyyy hh:mm:ss AM/PM"))
            
            
            
    'Stop
    For Each key In SysVars.Keys
    If InStr(1, dowebsite, key, vbTextCompare) <> 0 Then
        dowebsite = ReplaceVars(dowebsite, ("$" & key & "$"), SysVars(key))
    End If
    Next key
    
    
    
    Exit Function
errorocc:
    AppendLog Err.Description
End Function

Public Function showloginpage() As String
showloginpage = GetFile(frmAdmin.Rootpath.Text & "\login.htm")
End Function

Public Function printjavascript() As String
Dim p As String
p = "<SCRIPT>" & _
"if (document.location.host != '" & frmAdmin.LocalHostname.Text & "')" & _
"{" & _
 "document.location='http://" & frmAdmin.LocalHostname.Text & "';" & _
"}" & _
"</SCRIPT>"
printjavascript = p
End Function

Public Function GetFile(FileName As String) As String
'Stop
On Error GoTo ErrorOccured
Dim FileHandle0 As Long
Dim TempString As String
FileHandle0 = FreeFile
Open FileName For Input As FileHandle0
Do Until EOF(1)
Input #1, TempString
GetFile = GetFile & vbCrLf & TempString
Loop
Close #1
Exit Function
ErrorOccured:

GetFile = "<html>"
GetFile = GetFile & vbCrLf & "<head>"
GetFile = GetFile & vbCrLf & "<title>Untitled Document</title>"
GetFile = GetFile & vbCrLf & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
GetFile = GetFile & vbCrLf & "</head>"

GetFile = GetFile & vbCrLf & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
GetFile = GetFile & vbCrLf & "<div align=""center"">"
GetFile = GetFile & vbCrLf & "  <p><b><font face=""Arial, Helvetica, sans-serif"" color=""#666666"">***Application"
GetFile = GetFile & vbCrLf & "    error***<br>"
GetFile = GetFile & vbCrLf & "    An error occured in the webserver component of XCyteMail.</font></b></p>"
GetFile = GetFile & vbCrLf & "  <p>&nbsp;</p>"
GetFile = GetFile & vbCrLf & "  <p><b><font face=""Arial, Helvetica, sans-serif"" color=""#666666"">Module: " & "xcyte.webmail" & "<br>"
GetFile = GetFile & vbCrLf & "    <br>"
GetFile = GetFile & vbCrLf & "    Returned: " & Err.Description & "<br>"
GetFile = GetFile & vbCrLf & "    <br>"
GetFile = GetFile & vbCrLf & "    Error Number: " & Err.number & "</font></b></p>"
GetFile = GetFile & vbCrLf & "</div>"
GetFile = GetFile & vbCrLf & "</body>"
GetFile = GetFile & vbCrLf & "</html>"
End Function

Public Function AssemblePage(InputPage As String) As String
Dim TempPage As String
Dim TempString As String
Dim FileHandle0 As Long
Dim FileHandle1 As Long
FileHandle0 = FreeFile
FileHandle1 = FreeFile
On Error GoTo PageFault
Open frmAdmin.Rootpath.Text & "\header.htm" For Input As #FileHandle0
Do Until EOF(FileHandle0)
Input #FileHandle0, TempString

' Check to see if windowtext message is in there somewhere...
If InStr(1, TempString, "<i>#TEXT#</i>", vbTextCompare) <> 0 Then
    TempString = Replace(TempString, "<i>#TEXT#</i>", WindowTitle, 1, -1, vbTextCompare)
End If

If InStr(1, TempString, "<Replace>", vbTextCompare) <> 0 Then
    ' Found the cut point
    AssemblePage = AssemblePage & vbCrLf & InputPage
Else
    AssemblePage = AssemblePage & vbCrLf & TempString
End If
Loop
Close #FileHandle0
Exit Function
PageFault:

AssemblePage = "<html>"
AssemblePage = AssemblePage & vbCrLf & "<head>"
AssemblePage = AssemblePage & vbCrLf & "<title>Untitled Document</title>"
AssemblePage = AssemblePage & vbCrLf & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
AssemblePage = AssemblePage & vbCrLf & "</head>"

AssemblePage = AssemblePage & vbCrLf & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
AssemblePage = AssemblePage & vbCrLf & "<div align=""center"">"
AssemblePage = AssemblePage & vbCrLf & "  <p><b><font face=""Arial, Helvetica, sans-serif"" color=""#666666"">***Application"
AssemblePage = AssemblePage & vbCrLf & "    error***<br>"
AssemblePage = AssemblePage & vbCrLf & "    An error occured in the webserver component of XCyteMail.</font></b></p>"
AssemblePage = AssemblePage & vbCrLf & "  <p>&nbsp;</p>"
AssemblePage = AssemblePage & vbCrLf & "  <p><b><font face=""Arial, Helvetica, sans-serif"" color=""#666666"">Module: " & "xcyte.webmail" & "<br>"
AssemblePage = AssemblePage & vbCrLf & "    <br>"
AssemblePage = AssemblePage & vbCrLf & "    Returned: " & Err.Description & "<br>"
AssemblePage = AssemblePage & vbCrLf & "    <br>"
AssemblePage = AssemblePage & vbCrLf & "    Error Number: " & Err.number & "</font></b></p>"
AssemblePage = AssemblePage & vbCrLf & "</div>"
AssemblePage = AssemblePage & vbCrLf & "</body>"
AssemblePage = AssemblePage & vbCrLf & "</html>"
End Function

Public Function AssembleCompose() As String
Dim TempPage As String
Dim TempString As String
Dim FileHandle0 As Long
Dim FileHandle1 As Long
Dim FileHandle2 As Long
Dim LinkString As String
FileHandle0 = FreeFile
On Error GoTo PageFault
Open frmAdmin.Rootpath.Text & "\compose.htm" For Input As #FileHandle0
Do Until EOF(FileHandle0)
Input #FileHandle0, TempString

If InStr(1, TempString, "<ItemiseIndex>", vbTextCompare) <> 0 Then
    ' Found the cut point
    ' Now place our links in!
    If fso.FileExists(subfolder("Addressbooks") & "\" & CurrentUser & ".dat") = False Then
        FileHandle1 = FreeFile
        Open subfolder("Addressbooks") & "\" & CurrentUser & ".dat" For Append As #FileHandle1
        Close #FileHandle1
        LinkString = "<B>No entries found</B>"
        AssembleCompose = AssembleCompose & vbCrLf & LinkString
        GoTo OKGo
    End If
    FileHandle2 = FreeFile
    Dim sd As String
    Open subfolder("Addressbooks") & "\" & CurrentUser & ".dat" For Input As #FileHandle2
    Do Until EOF(FileHandle2)
    Input #FileHandle2, sd
    LinkString = "<a href=""""> " & sd & "</a>"
    AssembleCompose = AssembleCompose & vbCrLf & LinkString
    Loop
    Close #FileHandle2
Else
    AssembleCompose = AssembleCompose & vbCrLf & TempString
End If
OKGo:
Loop
Close #FileHandle0
Exit Function
PageFault:

AssembleCompose = "<html>"
AssembleCompose = AssembleCompose & vbCrLf & "<head>"
AssembleCompose = AssembleCompose & vbCrLf & "<title>Untitled Document</title>"
AssembleCompose = AssembleCompose & vbCrLf & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
AssembleCompose = AssembleCompose & vbCrLf & "</head>"

AssembleCompose = AssembleCompose & vbCrLf & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
AssembleCompose = AssembleCompose & vbCrLf & "<div align=""center"">"
AssembleCompose = AssembleCompose & vbCrLf & "  <p><b><font face=""Arial, Helvetica, sans-serif"" color=""#666666"">***Application"
AssembleCompose = AssembleCompose & vbCrLf & "    error***<br>"
AssembleCompose = AssembleCompose & vbCrLf & "    An error occured in the webserver component of XCyteMail.</font></b></p>"
AssembleCompose = AssembleCompose & vbCrLf & "  <p>&nbsp;</p>"
AssembleCompose = AssembleCompose & vbCrLf & "  <p><b><font face=""Arial, Helvetica, sans-serif"" color=""#666666"">Module: " & "xcyte.webmail" & "<br>"
AssembleCompose = AssembleCompose & vbCrLf & "    <br>"
AssembleCompose = AssembleCompose & vbCrLf & "    Returned: " & Err.Description & "<br>"
AssembleCompose = AssembleCompose & vbCrLf & "    <br>"
AssembleCompose = AssembleCompose & vbCrLf & "    Error Number: " & Err.number & "</font></b></p>"
AssembleCompose = AssembleCompose & vbCrLf & "</div>"
AssembleCompose = AssembleCompose & vbCrLf & "</body>"
AssembleCompose = AssembleCompose & vbCrLf & "</html>"
End Function

Public Function AssembleInbox(Src As String) As String
Dim TempPage As String
Dim TempString As String
Dim FileHandle0 As Long
Dim FileHandle1 As Long
Dim FileHandle2 As Long
Dim LinkString As String
FileHandle0 = FreeFile
On Error GoTo PageFault
Open frmAdmin.Rootpath.Text & "\inbox.htm" For Input As #FileHandle0
Do Until EOF(FileHandle0)
Input #FileHandle0, TempString

If InStr(1, TempString, "<MailList>", vbTextCompare) <> 0 Then
    ' Found the cut point
    ' Now place our links in!
    AssembleInbox = AssembleInbox & vbCrLf & Src
Else
    AssembleInbox = AssembleInbox & vbCrLf & TempString
End If
OKGo:
Loop
Close #FileHandle0
Exit Function
PageFault:

AssembleInbox = "<html>"
AssembleInbox = AssembleInbox & vbCrLf & "<head>"
AssembleInbox = AssembleInbox & vbCrLf & "<title>Untitled Document</title>"
AssembleInbox = AssembleInbox & vbCrLf & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
AssembleInbox = AssembleInbox & vbCrLf & "</head>"

AssembleInbox = AssembleInbox & vbCrLf & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
AssembleInbox = AssembleInbox & vbCrLf & "<div align=""center"">"
AssembleInbox = AssembleInbox & vbCrLf & "  <p><b><font face=""Arial, Helvetica, sans-serif"" color=""#666666"">***Application"
AssembleInbox = AssembleInbox & vbCrLf & "    error***<br>"
AssembleInbox = AssembleInbox & vbCrLf & "    An error occured in the webserver component of XCyteMail.</font></b></p>"
AssembleInbox = AssembleInbox & vbCrLf & "  <p>&nbsp;</p>"
AssembleInbox = AssembleInbox & vbCrLf & "  <p><b><font face=""Arial, Helvetica, sans-serif"" color=""#666666"">Module: " & "xcyte.webmail" & "<br>"
AssembleInbox = AssembleInbox & vbCrLf & "    <br>"
AssembleInbox = AssembleInbox & vbCrLf & "    Returned: " & Err.Description & "<br>"
AssembleInbox = AssembleInbox & vbCrLf & "    <br>"
AssembleInbox = AssembleInbox & vbCrLf & "    Error Number: " & Err.number & "</font></b></p>"
AssembleInbox = AssembleInbox & vbCrLf & "</div>"
AssembleInbox = AssembleInbox & vbCrLf & "</body>"
AssembleInbox = AssembleInbox & vbCrLf & "</html>"
End Function

Public Function GetUserfromIP(TargetUserIP As String) As String
If frmHTML.ListView1.ListItems.Count = 0 Then GetUserfromIP = "-1": Exit Function
For X = 1 To frmHTML.ListView1.ListItems.Count
If frmHTML.ListView1.ListItems(X).SubItems(2) = TargetUserIP Then
   GetUserfromIP = frmHTML.ListView1.ListItems(X).Text
   Exit Function
End If
Next X
' Hmm...not there!
' So let's just direct them to -1
GetUserfromIP = -1
End Function

Public Function GetIndexfromUser(TargetUser As String) As Integer
If frmHTML.ListView1.ListItems.Count = 0 Then GetIndexfromUser = "-1": Exit Function
For X = 1 To frmHTML.ListView1.ListItems.Count
If frmHTML.ListView1.ListItems(X).Text = TargetUser Then
   GetIndexfromUser = X
   Exit Function
End If
Next X
' Hmm...not there!
' So let's just send them -1
GetIndexfromUser = -1
End Function

Public Function AttachMailInsert(from As String, Subject As String, sDate As String, Size As String, MailID As String, FolderName As String)
          
Dim TempBuffer As String
          TempBuffer = "" & _
"                    <TR bgColor=#fff7e5>" & vbCrLf & _
"                      <TD name=""" & extractemail(from) & """><IMG alt=New" & vbCrLf & _
"                        height=18 hspace=5" & vbCrLf & _
"                        Src = ""i.newmail.gif""" & vbCrLf & _
"                      width=20></TD>" & vbCrLf & _
"                      <TD><INPUT name=" & MailID & " onclick=CCA(this);" & vbCrLf & _
"                        type=checkbox></TD>" & vbCrLf & _
"                      <TD>&nbsp;<A" & vbCrLf & _
"                        href = ""getmsg.asp?msg=" & MailID & "&folder=" & FolderName & """ > " & ExtractSimple(from) & vbCrLf & _
"                        </A>&nbsp;</TD>" & vbCrLf & _
"                      <TD>" & Subject & "&nbsp;</TD>" & vbCrLf & _
"                      <TD>" & Format(sDate, "dd mmm") & "</TD>" & vbCrLf & _
"                      <TD align=right>" & FormatSize(Size) & "k&nbsp;</TD></TR>"

          
AttachMailInsert = TempBuffer
End Function

Private Function ExtractSimple(FromString As String)
If InStr(1, FromString, "[", vbTextCompare) = 0 Then
    ExtractSimple = FromString
    Exit Function
End If
ExtractSimple = Left(FromString, InStr(1, FromString, "[", vbTextCompare) - 2)
End Function

Private Function ReplaceVars(InputData As String, Variable As String, ReplaceAs As String, Optional NotFound As String)
' Search through until all the vars are replaced
ReplaceVars = InputData
Dim Flag As Boolean
Do Until InStr(1, ReplaceVars, Variable, vbTextCompare) = 0
ReplaceVars = Replace(ReplaceVars, Variable, ReplaceAs, 1, -1, vbTextCompare)
Flag = True
Loop
End Function


Private Function MakeRedir(Redirurl As String) As String
Dim pjs As String
'// Redirect the user to the given url
pjs = "" & vbCrLf & _
"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2 Final//EN"">" & vbCrLf & _
"<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<title>Page redirector</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body>" & vbCrLf & _
"<SCRIPT>" & vbCrLf & _
"{" & vbCrLf & _
 "document.location='http://" & frmAdmin.LocalHostname.Text & "/cgi-bin/" & Redirurl & "';" & vbCrLf & _
"}" & vbCrLf & _
"</SCRIPT>" & vbCrLf & _
"</body>"
MakeRedir = pjs

End Function

Private Function LoadFile(InputFile As String) As String
Dim FileHandle0 As String
Dim temphandle As String
FileHandle0 = FreeFile
Open InputFile For Input As FileHandle0
Do Until EOF(FileHandle0)
Input #FileHandle0, temphandle
LoadFile = LoadFile & vbCrLf & temphandle
Loop
Close #FileHandle0
End Function

Private Function FormatSize(dSize As String) As String
    FormatSize = Val(dSize) / 1024
FormatSize = Round(FormatSize, 2)
End Function
