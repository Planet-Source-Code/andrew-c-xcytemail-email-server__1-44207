VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form POP3Session 
   Caption         =   "POP3 Applet"
   ClientHeight    =   1875
   ClientLeft      =   4365
   ClientTop       =   4170
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   0
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   2745
      Top             =   915
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "POP3Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public inbuf As String
Public outbuf As String

Public Timestamp As String
Public Username As String
Public Password As String
Public transactionstate As Boolean
Public deletedmsgs As New Dictionary

Private Sub ws_Close()
ws.Close '
Me.Hide
Unload Me
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
'Stop
    Dim q As String, ts As TextStream
    ws.getdata q
    DoEvents
    If ControlLogging = True Then AppendLog "POP3>" & q
    inbuf = inbuf & q
    If InStr(1, inbuf, vbCrLf) = 0 Then Exit Sub
    On Error Resume Next
    K = Mid(inbuf, 1, InStr(1, inbuf, " ") - 1)
    v = Mid(inbuf, InStr(1, inbuf, " ") + 1)
    inbuf = Mid(inbuf, InStr(1, inbuf, vbCrLf) + 2)
    If Right(v, 2) = vbCrLf Then v = Mid(v, 1, Len(v) - 2)
    If K = "" And v <> "" Then K = v: v = ""
    On Error GoTo 0
    Select Case transactionstate
    Case False
        Select Case UCase(K)
        Case "QUIT"
            outbuf = outbuf & "+OK Please come again!" & vbCrLf
            ws.Close
            Me.Hide
            Unload Me
            Exit Sub
        Case "USER"
            Username = v
            If AccountExists(Username) = True Then
                outbuf = outbuf & "+OK Username accepted!" & vbCrLf
            Else
                outbuf = outbuf & "-ERR Username rejected!" & vbCrLf
            End If
        Case "PASS"
            'MsgBox GetEnabledState(Username)
            If GetEnabledState(Username) = False Then
                outbuf = outbuf & "-ERR Account has been disabled by an administrator!" & vbCrLf
                GoTo skipvalidate
            End If
            If ValidatePassword(Username, CStr(v)) = True Then
                Password = v
                transactionstate = True
                outbuf = outbuf & "+OK Login completed!" & vbCrLf
                File1.Path = subfolder(Username)
            Else
                outbuf = outbuf & "-ERR Login rejected, please try again!" & vbCrLf
                Username = ""
                Password = ""
            End If
skipvalidate:
        Case "APOP"
            digest = Mid(v, InStr(1, v, " ") + 1)
            v = Mid(v, 1, InStr(1, v, " ") - 1)
            Username = v
            If LCase(digest) = LCase(checksum(Timestamp & getaccountinfo(Username, "pw"))) Then
                transactionstate = True
                outbuf = outbuf & "+OK Secure login established" & vbCrLf
            Else
                outbuf = outbud & "-ERR Error in secure login" & vbCrLf
            End If
        Case Else
            outbuf = outbuf & "-ERR Bad command at authenticate stage" & vbCrLf
        End Select
    Case True
        Select Case UCase(K)
        Case "NOOP"
            outbuf = outbuf & "+OK Replied to ping pong." & vbCrLf
        Case "DELE"
            deletedmsgs(v) = True
            outbuf = outbuf & "+OK Messages will be cleared on logout or RSET" & vbCrLf
        Case "RSET"
            deletedmsgs.RemoveAll
            outbuf = outbuf & "+OK All messages have been deleted!" & vbCrLf
        Case "UIDL"
            DoEvents
            File1.Refresh
            If v = "" Then
                outbuf = outbuf & "+OK" & vbCrLf
                For a = 1 To File1.ListCount - 1
                    outbuf = outbuf & a & " " & Mid(File1.List(a), 1, Len(File1.List(a)) - 4) & vbCrLf
                Next a
                outbuf = outbuf & "." & vbCrLf
            Else
                outbuf = outbuf & "+OK " & v & " " & Mid(File1.List(v), 1, Len(File1.List(v)) - 4) & vbCrLf
            End If
        Case "LIST"
        DoEvents
        File1.Refresh
            If v = "" Then
                outbuf = outbuf & "+OK" & vbCrLf
                For a = 1 To File1.ListCount - 1
                    outbuf = outbuf & a & " " & getmailsize(fso.BuildPath(File1.Path, File1.List(a))) & vbCrLf
                Next a
                outbuf = outbuf & "." & vbCrLf
            Else
                outbuf = outbuf & "+OK " & v & " " & getmailsize(fso.BuildPath(File1.Path, File1.List(v))) & vbCrLf
            End If
        Case "STAT"
'           Stop
            DoEvents
            File1.Refresh
            Sum = 0
            For a = 0 To File1.ListCount - 1
                Sum = Sum + getmailsize(fso.BuildPath(File1.Path, File1.List(a)))
            Next a
            outbuf = outbuf & "+OK " & File1.ListCount & " " & Sum & vbCrLf
        Case "RETR"
            If v = "" Then
                outbuf = "+ERR Bad RETR syntax"
                Exit Sub
            End If
            File1.Refresh
            Set ts = fso.OpenTextFile(fso.BuildPath(File1.Path, File1.List(v)))
            ts.SkipLine
            ts.SkipLine
            ts.SkipLine
            ts.SkipLine
            outbuf = outbuf & "+OK" & vbCrLf & ts.ReadAll & vbCrLf & "." & vbCrLf
            ts.Close
        Case "TOP"
           File1.Refresh
            numlines = Mid(v, InStr(1, v, " ") + 1)
            v = Mid(v, 1, InStr(1, v, " ") - 1)
            Set ts = fso.OpenTextFile(fso.BuildPath(File1.Path, File1.List(v)))
            ts.SkipLine
            ts.SkipLine
            ts.SkipLine
            ts.SkipLine
            outbuf = outbuf & "+OK" & vbCrLf
keepgoing:
            z = ts.ReadLine
            If z <> "" Then outbuf = outbuf & z & vbCrLf: GoTo keepgoing
            outbuf = outbuf & vbCrLf
            On Error Resume Next
            For a = 1 To numlines
                outbuf = outbuf & ts.ReadLine & vbCrLf
            Next a
            On Error GoTo 0
            outbuf = outbuf & "." & vbCrLf
            ts.Close
        Case "QUIT"
            File1.Refresh
            outbuf = outbuf & "+OK Please come again!" & vbCrLf
            For Each a In deletedmsgs.keys
                fso.DeleteFile fso.BuildPath(File1.Path, File1.List(a)), True
                'RemoveEmailIndex Left(File1.List(a), Len(File1.List(a)) - 4)
            Next
            ws.Close
            Me.Hide
            Unload Me
        Case Else
            outbuf = outbuf & "-ERR Unknown command!" & vbCrLf
        End Select
    End Select
    On Error Resume Next
    ws.SendData outbuf
    outbuf = ""
End Sub

Private Function getmailsize(FileName As String) As Long
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

Private Sub ws_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ws.Close
    Me.Hide
    Unload Me
End Sub

Private Function previewtopline() As String
    a = InStr(1, inbuf, vbCrLf)
    If a > 0 Then
        previewtopline = Mid(inbuf, 1, a - 1)
    End If
End Function

Private Function pulltopline() As String
    a = InStr(1, inbuf, vbCrLf)
    If a > 0 Then
        pulltop = Mid(inbuf, 1, a - 1)
        inbuf = Mid(inbuf, a + 2)
    End If
End Function

