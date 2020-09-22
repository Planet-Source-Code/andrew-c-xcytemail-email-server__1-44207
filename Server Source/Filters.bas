Attribute VB_Name = "Filters"
Dim Void As String
Dim comstring As String
Dim PublicMSGVal As Boolean
Option Explicit

' Some info on email filtering
' General email file structure:
' Line 1: From Address
' Line 2: To Address
' Line 3: "0"
' Line 4: Date/Time
' Line 5: Message Data

Public Function FilterMessage(MessageFN As String) As Boolean
Dim FilterType As String
Dim FilterInfo As String
Dim BaseString As String
Dim FileHandle1 As Long
FileHandle1 = FreeFile
Close #FileHandle1
'On Error GoTo IncorrectFile
Open App.Path & "\Config\Filters.cfg" For Input As #FileHandle1
Do Until EOF(FileHandle1)
Input #FileHandle1, Void
Input #FileHandle1, Void
Input #FileHandle1, FilterType
Input #FileHandle1, FilterInfo
Input #FileHandle1, Void

BaseString = Left(FilterInfo, InStr(1, FilterInfo, ":", vbTextCompare) - 1)
comstring = Right(FilterInfo, Len(FilterInfo) - Len(BaseString) - 2)
FilterType = Right(FilterType, Len(FilterType) - InStr(1, FilterType, ":", vbTextCompare) - 1)
Stop
Select Case FilterType
Case "Sender"
    RunSenderFilter MessageFN, comstring
Case "Recipient"
    RunToFilter MessageFN, comstring
Case "Subject"
    RunSubjectFilter MessageFN, comstring
Case "Keywords"
    RunKeywordFilter MessageFN, comstring
End Select
Loop
Close #FileHandle1
DoEvents
FilterMessage = PublicMSGVal
Exit Function
IncorrectFile:
AppendLog "***Incorrect filter.cfg file format!"
End Function

Private Sub RunSenderFilter(MessageFN As String, ParString As String)
' Note about sender type filter
' Only one email per filter
Dim fTemp As String
Dim cFilter As String
Dim fData As String
Dim EndFilename As String
Dim FileHandle0 As Long
Dim FileHandle1 As Long
Dim FileHandle2 As Long
Dim Filehandle3 As Long
Dim filehandle4 As Long

FileHandle0 = FreeFile
Open MessageFN For Input As #FileHandle0
Line Input #FileHandle0, fData
Close #FileHandle0

' Check for more than 1 email in list
If InStr(1, ParString, ";", vbTextCompare) = 0 Then
   '1 Email in list
    If InStr(1, LCase(fData), LCase(ParString), vbTextCompare) <> 0 Then
        ' Sender if to be filtered
        ' Move the message data
        EndFilename = Right(MessageFN, Len(MessageFN) - InStrRev(MessageFN, "\", -1, vbTextCompare))
        FileHandle1 = FreeFile
        Open MessageFN For Input As #FileHandle1
        FileHandle2 = FreeFile
        Open subfolder("Filtered") & "\" & EndFilename For Output As #FileHandle2
        Do Until EOF(FileHandle1)
            Line Input #FileHandle1, cFilter
            Print #FileHandle2, cFilter
        Loop
        Close #FileHandle1
        Close #FileHandle2
        Kill MessageFN
        PublicMSGVal = True
        ' Found the banned email address, so quit the loop
        Exit Sub
    End If
End If
Close #FileHandle1
Close #FileHandle2
Do Until InStr(1, ParString, ";", vbTextCompare) = 0
    ' Get the first(or only) email in the list
    fTemp = Left(ParString, InStr(1, ParString, ";", vbTextCompare) - 1)
    ' See if the listed email is the sender
    If InStr(1, LCase(fData), LCase(fTemp), vbTextCompare) <> 0 Then
        ' Sender if to be filtered
        ' Move the message data
        EndFilename = Right(MessageFN, Len(MessageFN) - InStrRev(MessageFN, "\", -1, vbTextCompare))
        Filehandle3 = FreeFile
        Open MessageFN For Input As #Filehandle3
        filehandle4 = FreeFile
        Open subfolder("Filtered") & "\" & EndFilename For Output As #filehandle4
        Do Until EOF(Filehandle3)
            Line Input #Filehandle3, cFilter
            Print #filehandle4, cFilter
        Loop
        Close #Filehandle3
        Close #filehandle4
        Kill MessageFN
        ' Found the banned email address, so quit the loop
        PublicMSGVal = True
        Exit Sub
    End If
    ' No match, so trim off the email address
    ParString = Right(ParString, Len(ParString) - Len(fTemp) - 2)
Loop
End Sub

Private Sub RunToFilter(MessageFN As String, ParString As String)
' Note about sender type filter
' Only one email per filter
Dim fTemp As String
Dim cFilter As String
Dim fData As String
Dim SenderData As String
Dim EndFilename As String
Dim FileHandle0 As Long
Dim FileHandle1 As Long
Dim FileHandle2 As Long
Dim Filehandle3 As Long
Dim filehandle4 As Long
Dim filehandle5 As Long
FileHandle0 = FreeFile
Open MessageFN For Input As #FileHandle0
Line Input #FileHandle0, SenderData
Line Input #FileHandle0, fData
Close #FileHandle0

' Check for more than 1 email in list
If InStr(1, ParString, ";", vbTextCompare) = 0 Then
   '1 Email in list
    If InStr(1, LCase(fData), LCase(ParString), vbTextCompare) <> 0 Then
        ' Sender if to be filtered
        ' Move the message data
        EndFilename = Right(MessageFN, Len(MessageFN) - InStrRev(MessageFN, "\", -1, vbTextCompare))
        FileHandle1 = FreeFile
        Open MessageFN For Input As #FileHandle1
        FileHandle2 = FreeFile
        Open subfolder("Filtered") & "\" & EndFilename For Output As #FileHandle2
        Do Until EOF(FileHandle1)
            Line Input #FileHandle1, cFilter
            Print #FileHandle2, cFilter
        Loop
        Close #FileHandle1
        Close #FileHandle2
        PublicMSGVal = True
        Kill MessageFN
        ' Message has been filtered. Send an email to
        ' sender!
        Dim ts As TextStream
        Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
        ts.WriteLine "Postman@" & frmAdmin.LocalHostname.Text
        ts.WriteLine SenderData
        ts.WriteLine "0"
        ts.WriteLine Now
        ts.WriteLine "from: Postman@" & frmAdmin.LocalHostname.Text
        ts.WriteLine "to: " & SenderData
        ts.WriteLine "subject: Your email has been filtered!"
        ts.WriteBlankLines 2
        ts.WriteLine "This is an automated message"
        ts.WriteLine "Do not respond to this email."
        ts.WriteBlankLines 2
        ts.WriteLine "XCyteMail server automatically filtered your email."
        ts.WriteLine "The email address: '" & fData & "' was automatically filtered!"
        ts.Close
        Exit Sub
    End If
End If
Close #FileHandle1
Close #FileHandle2
Do Until InStr(1, ParString, ";", vbTextCompare) = 0
    ' Get the first(or only) email in the list
    fTemp = Left(ParString, InStr(1, ParString, ";", vbTextCompare) - 1)
    ' See if the listed email is the sender
    If InStr(1, LCase(fData), LCase(fTemp), vbTextCompare) <> 0 Then
        ' Sender if to be filtered
        ' Move the message data
        EndFilename = Right(MessageFN, Len(MessageFN) - InStrRev(MessageFN, "\", -1, vbTextCompare))
        Filehandle3 = FreeFile
        Open MessageFN For Input As #Filehandle3
        filehandle4 = FreeFile
        Open subfolder("Filtered") & "\" & EndFilename For Output As #filehandle4
        Do Until EOF(Filehandle3)
            Line Input #Filehandle3, cFilter
            Print #filehandle4, cFilter
        Loop
        Close #Filehandle3
        Close #filehandle4
        PublicMSGVal = True
        Kill MessageFN
        ' Found the banned email address, so
        ' send an email to the sender and quit!
        Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
        ts.WriteLine "Postman@" & frmAdmin.LocalHostname.Text
        ts.WriteLine SenderData
        ts.WriteLine "0"
        ts.WriteLine Now
        ts.WriteLine "from: Postman@" & frmAdmin.LocalHostname.Text
        ts.WriteLine "to: " & SenderData
        ts.WriteLine "subject: Your email has been filtered!"
        ts.WriteBlankLines 2
        ts.WriteLine "This is an automated message"
        ts.WriteLine "Do not respond to this email."
        ts.WriteBlankLines 2
        ts.WriteLine "XCyteMail server automatically filtered your email."
        ts.WriteLine "The email address: '" & fData & "' was automatically filtered!"
        ts.Close
        Exit Sub
    End If
    ' No match, so trim off the email address
    ParString = Right(ParString, Len(ParString) - Len(fTemp) - 2)
Loop
End Sub


Private Sub RunKeywordFilter(MessageFN As String, ParString As String)

End Sub

Private Sub RunSubjectFilter(MessageFN As String, ParString As String)

End Sub
