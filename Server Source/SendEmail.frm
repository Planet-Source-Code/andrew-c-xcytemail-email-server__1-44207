VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form SendEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SendMail Applet"
   ClientHeight    =   2070
   ClientLeft      =   3780
   ClientTop       =   3945
   ClientWidth     =   5085
   Icon            =   "SendEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock ws 
      Left            =   4440
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label txt 
      Caption         =   "Label1"
      Height          =   1170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Shape pb 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   320
      Left            =   0
      Top             =   1218
      Width           =   4680
   End
End
Attribute VB_Name = "SendEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private afrom As String
Private ato As String
Private Data As String
Private attemptcount As Long
Private datafilename As String
Public MailPending As Boolean
Private stage As Long

Public Sub setup(FileName As String)
'Stop
On Error GoTo errorocc
MailPending = True
'AppendLog "MailPending = True"
If ActiveLogging = True Then
    ErrorLog.WriteError vbCrLf
    ErrorLog.WriteError "Preparing outgoing email: " & FileName
End If
Dim sString As String
Dim OpenBuffer As Long
'frmMain.Timer1.Enabled = False
On Error Resume Next
    If Not fso.FileExists(FileName) Then
        MailPending = False
        Unload Me
        Exit Sub
    End If
    datafilename = FileName
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(FileName)
    afrom = ts.ReadLine
    ato = ts.ReadLine
    attemptcount = Val(ts.ReadLine) + 1
    donttrytill = ts.ReadLine
    Data = ts.ReadAll
    ts.Close
    
    If ActiveLogging = True Then
        ErrorLog.WriteError "From: " & afrom
        ErrorLog.WriteError "To: " & ato
        ErrorLog.WriteError "Attemptcount: " & attemptcount
    End If
    
'    If CDate(donttrytill) > Now Then
'        Me.Hide
'        Unload Me
'        Exit Sub
'    End If
    
    If attemptcount = 8 Then
        Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
        ts.WriteLine "Postman@" & frmAdmin.LocalHostname.Text
        ts.WriteLine afrom
        ts.WriteLine "0"
        ts.WriteLine Now
        ts.WriteLine "from: Postman@" & frmAdmin.LocalHostname.Text
        ts.WriteLine "to: " & afrom
        ts.WriteLine "subject: Message delivery failure!"
        ts.WriteBlankLines 2
        ts.WriteLine "This is an automated email"
        ts.WriteLine "Do not respond to this message!"
        ts.WriteBlankLines 2
        ts.WriteLine "Your email could not be sent to the intended recipient!"
        ts.WriteLine "The mail server will automatically resend your message."
        ts.Close
        ErrorLog.WriteError "Message delivery failure."
        ErrorLog.WriteError vbCrLf
        attemptcount = attemptcount + 1
    End If
    
    If attemptcount = 12 Then
        Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
        ts.WriteLine "Postman@" & frmAdmin.LocalHostname.Text
        ts.WriteLine afrom
        ts.WriteLine "0"
        ts.WriteLine Now
        ts.WriteLine "from: Postman@" & frmAdmin.LocalHostname.Text
        ts.WriteLine "to: " & afrom
        ts.WriteLine "subject: Failed in sending your message"
        ts.WriteBlankLines 2
'        Stop
        ts.WriteLine "This is an automated email"
        ts.WriteLine "Do not respond to this message!"
        ts.WriteBlankLines 2
        ts.WriteLine "Previous attempts to send your message failed!"
        ts.WriteLine "Your message has been removed from the system."
        ts.Close
        fso.DeleteFile FileName
        ErrorLog.WriteError "Message delivery failure."
        ErrorLog.WriteError "Message removed from system."
        ErrorLog.WriteError vbCrLf
        MailPending = False
        attemptcount = 0
        'Exit Sub
    End If
    Set ts = fso.OpenTextFile(FileName, ForWriting, False)
    ts.WriteLine afrom
    ts.WriteLine ato
    ts.WriteLine attemptcount
    ts.WriteLine Now + (2 ^ (attemptcount - 1)) / 1440
    ts.Write Data
    ts.Close
    If mxlookup(Mid(ato, InStr(1, ato, "@") + 1)) = "" Then
        Debug.Print "Failed sending!"
        Debug.Print "Tried " & (Mid(ato, InStr(1, ato, "@") + 1))
        Debug.Print "Got of mx: " & mxlookup(Mid(ato, InStr(1, ato, "@") + 1))
        If attemptcount = 1 Then
            Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
            ts.WriteLine "Postman@" & frmAdmin.LocalHostname.Text
            ts.WriteLine afrom
            ts.WriteLine "0"
            ts.WriteLine Now
            ts.WriteLine "from: Postman@" & frmAdmin.LocalHostname.Text
            ts.WriteLine "to: " & afrom
            ts.WriteLine "subject: Unable to deliverer your message"
            ts.WriteBlankLines 2
            ts.WriteLine "This is an automated message"
            ts.WriteLine "Do not respond to this email."
            ts.WriteBlankLines 2
            ts.WriteLine "The server was unable to locate the host: " & Mid(ato, InStr(1, ato, "@") + 1)
            ts.WriteLine "Your mail message could not be delivered!"
            ts.WriteLine "The server will automatically retry again 7 times over the next 4 hours, and notify you if your message still remains undelivered. If it is still undelivered at this date, the server will then try another 4 times over the next 4 days. After 4 days, if your message still remains undelivered, you will be notified, and the email will be removed from the system."
            ts.Close
            ErrorLog.WriteError "Unable to locate remote mail server."
         '   ErrorLog.WriteError vbCrLf
            attemptcount = attemptcount + 1
        
        End If
        Me.Hide
        Unload Me
        SendEmail.MailPending = False
        frmMain.File1.Refresh
        Exit Sub
    End If
'    Stop
    txt = "From: " & afrom & vbNewLine & "To: " & ato & vbNewLine & "Size: " & Len(Data)
    pb.Width = 0
    'Me.Show
    ws.Close
    
    If InStr(1, ato, "@" & frmAdmin.LocalHostname.Text) Then
        ' Trying to email ourselves.
        Dim sendit As New inmail
        sendit.moreincomming "HELO " & frmAdmin.LocalHostname.Text & " webmail" & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "MAIL FROM: " & afrom & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "RCPT TO: " & ato & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "DATA" & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming Data & vbCrLf & "." & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "QUIT" & vbCrLf
        sendit.parsebuffer
        fso.DeleteFile FileName
        Me.Hide
        Unload Me
        MailPending = False
        Exit Sub
    End If
'    Stop
    ws.connect mxlookup(Mid(ato, InStr(1, ato, "@") + 1)), 25
'    ws.SendData "HELO " & ws.LocalIP & vbCrLf

Exit Sub
errorocc:
MsgBox Err.Description
End Sub

Private Sub ws_Connect()
    stage = 0
    Data = Data & vbCrLf & "." & vbCrLf
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Sleep 1000
    DoEvents
    Dim a As String
    On Error Resume Next
    ws.getdata a
    If ControlLogging = True Then AppendLog "SMTP>" & a
    On Error GoTo 0
    If Err.Description <> "" Then a = "220 connect ok" & vbCrLf
    If Mid(a, 1, 3) = "250" Or Mid(a, 1, 3) = "220" Or Mid(a, 1, 3) = "351" Or Mid(a, 1, 3) = "354" Then
        stage = stage + 1
        pb.Width = stage * 0.1
    Else
        If attemptcount = 1 Then
            Dim ts As TextStream
            Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
            ts.WriteLine "Postman@" & frmAdmin.LocalHostname.Text
            ts.WriteLine afrom
            ts.WriteLine "0"
            ts.WriteLine Now
            ts.WriteLine "from: Postman@" & frmAdmin.LocalHostname.Text
            ts.WriteLine "to: " & afrom
            ts.WriteLine "subject: Unable to deliver message."
            ts.WriteBlankLines 2
            ts.WriteLine "This is an automated message"
            ts.WriteLine "Do not respond to this email."
            ts.WriteBlankLines 2
            ts.WriteLine "Your message could not be delivered!"
            ts.WriteLine "The server " & ws.RemoteHost & " returned the following line:"
            ts.WriteLine ""
            ts.WriteLine a
            ts.WriteLine ""
            ts.WriteLine "The server will try again 7 times over the next 4 hours, and notify you if your message still remains undelivered. If it is still undelivered at this date, the server will then try another 4 times over the next 4 days. After 4 days, if your message still remains undelivered, you will be notified, and the email will be removed from the system."
            ts.Close
        End If
        Hide
        Unload Me
        MailPending = False
        'AppendLog "MailPending = false"
        frmMain.File1.Refresh
        Exit Sub
    End If
    
    Debug.Print stage & ":" & a
    Select Case stage
    Case 1
        ws.SendData "HELO " & frmAdmin.LocalHostname.Text & vbCrLf
    Case 2
        ws.SendData "MAIL FROM: " & afrom & vbCrLf
    Case 3
        ws.SendData "RCPT TO: " & ato & vbCrLf
    Case 4
        ws.SendData "DATA" & vbCrLf
    Case 5
        While Len(Data) > 1
            ws.SendData Mid(Data, 1, InStr(1, Data, vbCrLf) + 1)
            Data = Mid(Data, InStr(1, Data, vbCrLf) + 2)
        Wend
    Case 6
        'Stop
        Sleep 500
        ws.SendData "QUIT" & vbCrLf
        On Error Resume Next
        fso.DeleteFile datafilename, True
        Me.Hide
        MailPending = False
        'AppendLog "MailPending = false"
        frmMain.File1.Refresh
        Unload Me
    End Select
End Sub

Private Sub ws_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If attemptcount = 1 Then
            Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
            ts.WriteLine "Postman@" & frmAdmin.LocalHostname.Text
            ts.WriteLine afrom
            ts.WriteLine "0"
            ts.WriteLine Now
            ts.WriteLine "from: Postman@" & frmAdmin.LocalHostname.Text
            ts.WriteLine "to: " & afrom
            ts.WriteLine "subject: Unable to deliverer your message"
            ts.WriteBlankLines 2
            ts.WriteLine "This is an automated message"
            ts.WriteLine "Do not respond to this email."
            ts.WriteBlankLines 2
            ts.WriteLine "The server was unable to locate the host: " & Mid(ato, InStr(1, ato, "@") + 1)
            ts.WriteLine "Your mail message could not be delivered!"
            ts.WriteLine "The server will automatically retry again 7 times over the next 4 hours, and notify you if your message still remains undelivered. If it is still undelivered at this date, the server will then try another 4 times over the next 4 days. After 4 days, if your message still remains undelivered, you will be notified, and the email will be removed from the system."
            ts.Close
            ErrorLog.WriteError "Message delivery failure."
            ErrorLog.WriteError vbCrLf
End If
        Me.Hide
        Unload Me
        MailPending = False
        Exit Sub
    ws.Close
End Sub

