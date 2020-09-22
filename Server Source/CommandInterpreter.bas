Attribute VB_Name = "CommandInterpreter"
Option Explicit

Public Sub Processcommand(sRawCommand As String)
DoEvents
'Stop
Dim CommandRoot As String
Dim Parameters() As String 'Max 10 parameters
Dim ParameterCount As Integer
Dim LogFreeFile As Long
ParameterCount = 0
If keylogger = True Then
   LogFreeFile = FreeFile
   Open LogStruct For Append As LogFreeFile
   Print #LogFreeFile, sRawCommand
   Close #1
End If
If Editing = True Then
   If sRawCommand = "q" Then
      Editing = False
      AppendLog ">"
   End If
   If sRawCommand = "w" Then
      Open CurrentFile For Output As #1
      Print #1, Editor.FileData
      Close #1
   End If
   Editor.FileData = Editor.FileData & vbCrLf & sRawCommand
   Exit Sub
End If

' Check for static commands
If LCase(sRawCommand) = "resetclose" Then
      RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "NOSHOWDIALOG", 0
      RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "Software\XcyteServer\Settings", "NOSHOWMODE", 0
      AppendLog "Resetting close state variable"
      AppendOK
End If


' Check for multi-parameter commands
If InStr(1, sRawCommand, " ", vbTextCompare) = 0 Then
   GoTo HandleSingles
End If

' Seperate the command root and parameters
CommandRoot = Trim(LCase(Left(sRawCommand, InStr(1, sRawCommand, " ", vbTextCompare) - 1)))
sRawCommand = Right(sRawCommand, Len(sRawCommand) - Len(CommandRoot) - 1)

' Filter parameters into boxes
Do While InStr(1, sRawCommand, " ", vbTextCompare) <> 0
ReDim Preserve Parameters(ParameterCount)
Parameters(ParameterCount) = Left(sRawCommand, InStr(1, sRawCommand, " ", vbTextCompare) - 1)
sRawCommand = Right(sRawCommand, Len(sRawCommand) - Len(Parameters(ParameterCount)) - 1)
If sRawCommand = " " Then sRawCommand = ""
ParameterCount = ParameterCount + 1
Loop

' Filter the last command into the box
ReDim Preserve Parameters(ParameterCount)
Parameters(ParameterCount) = sRawCommand

' Place multi-parameter commands here.
' For example: StartService POP3
If CommandRoot = "startservice" Then
   Dim serviceToggle As Boolean
   'StartService command
   If Parameters(0) = "activelog" Then ActiveLogging = True: serviceToggle = True: AppendLog "Active logging enabled": ActiveLogging = True
   If Parameters(0) = "controllog" Then ControlLogging = True: serviceToggle = True: AppendLog "Control logging enabled": ControlLogging = True
   If Parameters(0) = "pop3" Then StartService ("POP3"): serviceToggle = True
   If Parameters(0) = "smtp" Then StartService ("SMTP"): serviceToggle = True
   If Parameters(0) = "webmail" Then StartService ("WEBMAIL"): serviceToggle = True
   If Parameters(0) = "keylog" Then StartService ("KEYLOG"): serviceToggle = True: AppendLog "Key logger enabled": keylogger = True
   If Parameters(0) = "systray" Then StartService ("SYSTRAY"): serviceToggle = True
       
       
If serviceToggle = False Then AppendLog "   *Unrecognised service " & Parameters(0)
End If

If CommandRoot = "stopservice" Then
   Dim serviceToggle2 As Boolean
   'StopService command
   If Parameters(0) = "controllog" Then ControlLogging = False: serviceToggle = True: AppendLog "Control logging disabled": ControlLogging = False
   If Parameters(0) = "activelog" Then ActiveLogging = False: serviceToggle = True: AppendLog "Active logging disabled": ActiveLogging = False: frmMain.SvrLog.Text = frmMain.SvrLog.Text & vbCrLf & vbCrLf & "You have stopped the activelog service!" & vbCrLf & "Without active logging enabled, the console will not function." & vbCrLf & "Press F6 to start active logging again..." & vbCrLf
   If Parameters(0) = "pop3" Then StopService ("POP3"): serviceToggle = True
   If Parameters(0) = "smtp" Then StopService ("SMTP"): serviceToggle = True
   If Parameters(0) = "webmail" Then StopService ("WEBMAIL"): serviceToggle = True
   If Parameters(0) = "keylog" Then StopService ("KEYLOG"): serviceToggle = True: AppendLog "Key logger disabled": keylogger = False

If serviceToggle = False Then AppendLog "   *Unrecognised service " & Parameters(0)
End If

If CommandRoot = "restartservice" Then
   Dim serviceToggle3 As Boolean
   'StopService command
   If Parameters(0) = "activelog" Then ActiveLogging = False: serviceToggle = True
   If Parameters(0) = "pop3" Then StopService ("POP3"): serviceToggle = True
   If Parameters(0) = "smtp" Then StopService ("SMTP"): serviceToggle = True
   If Parameters(0) = "webmail" Then StopService ("WEBMAIL"): serviceToggle = True
   If Parameters(0) = "keylog" Then StopService ("KEYLOG"): serviceToggle = True

   If Parameters(0) = "activelog" Then ActiveLogging = True: serviceToggle = True
   If Parameters(0) = "pop3" Then StartService ("POP3"): serviceToggle = True
   If Parameters(0) = "smtp" Then StartService ("SMTP"): serviceToggle = True
   If Parameters(0) = "webmail" Then StartService ("WEBMAIL"): serviceToggle = True
   If Parameters(0) = "keylog" Then StartService ("KEYLOG"): serviceToggle = True
   If Parameters(0) = "all" Then
      AppendLog "Restarting all services..."
      StartSilent = True
      StopService "POP3"
      StartSilent = True
      StopService "SMTP"
      StartSilent = True
      StopService "WEBMAIL"
      StartSilent = True
      StartService "POP3"
      StartSilent = True
      StartService "SMTP"
      StartSilent = True
      StartService "WEBMAIL"
      AppendOK
      serviceToggle = True
      Exit Sub
  End If
If serviceToggle = False Then AppendLog "   *Bad command syntax: " & Parameters(0)
End If

If CommandRoot = "help" Then
   HandleHelpSubSystem Parameters(0)
End If

If CommandRoot = "getmx" Or CommandRoot = "mxlookup" Then
   Dim MXValue As String
   AppendLog "Performing MX lookup. Please wait..."
   MXValue = mxlookup(Parameters(0))
   DoEvents
   If MXValue = "" Or MXValue = " " Then AppendLog "***Unable to complete lookup! Not connected to internet.": AppendLog ">": Exit Sub
   AppendLog "MX query returned: " & MXValue
End If

If CommandRoot = "edit" Then
   Dim sFilesize As Integer
   ' start the editor
   If FileExists(Parameters(0)) = False Then
      sFilesize = "0"
   Else
      sFilesize = FileLen(Parameters(0))
   End If
   Editing = True
   CurrentFile = Parameters(0)
   AppendLog "---Text Editor---"
   AppendLog "Current file: " & CurrentFile
   AppendLog "Filesize: " & FileLen(Parameters(0)) & " Bytes"
   AppendLog " "
   If FileExists(Parameters(0)) = False Then
      AppendLog " "
      Exit Sub
   Else
      Dim LineData As String
      Open Parameters(0) For Input As #1
      Do Until EOF(1)
      Line Input #1, LineData
      Editor.FileData = Editor.FileData & vbCrLf & LineData
      Loop
      Close #1
      AppendLog Editor.FileData
      AppendLog vbCrLf
      Exit Sub
   End If
End If
If CommandRoot = "shutdown" Then
   DoShutdownTimer Parameters(0)
   Exit Sub
End If
AppendLog ">"
Exit Sub

HandleSingles:
' Place single parameter commands here.
' For example: ShutdownServer or Kickall.
sRawCommand = LCase(sRawCommand)
If sRawCommand = "clear" Then
   frmMain.SvrLog.Text = ""
End If
If sRawCommand = "listservices" Then
    AppendLog " "
    AppendLog "Available services:"
    AppendLog "---------------------------------------------------"
    AppendLog "ActiveLog        - Active logging"
    AppendLog "ControlLog       - Active data logging"
    AppendLog "POP3             - Manages the pop3 protocol"
    AppendLog "SMTP             - Manages the smtp protocol"
    AppendLog "WEBMAIL          - Manages the webmail interface"
    AppendLog "KEYLOG           - Handles the console key logger"
    AppendLog "SYSTAY           - Manages system tray support"
    AppendLog "SideBarAutoHide  - Autohides the sidebar object"
    AppendLog "---------------------------------------------------"
    AppendLog " "
End If

If sRawCommand = "listcommands" Or sRawCommand = "list" Then
    AppendLog " "
    AppendLog "Available commands:"
    AppendLog "-------------------------------------------------------------------------"
    AppendLog "StartService [Servicename]    - Starts the given service"
    AppendLog "StopService  [Servicename]    - Stops the given service"
    AppendLog "RestartService [ServiceName]  - Restarts the given service"
    AppendLog "RestartService all            - Systematically restarts all services"
    AppendLog "ShutdownServer                - Terminates the mail server software"
    AppendLog "Shutdown [Time]               - Terminates the server after the given time"
    AppendLog "ListServices                  - Displays all the startable services"
    AppendLog "Clear                         - Clears the console output window"
    AppendLog "GetMX [Domain]                - Performs an MX lookup on the domain"
    AppendLog "ResetClose                    - Resets the close action variable"
    AppendLog "Help [String]                 - Displays the associated help file"
    AppendLog "-------------------------------------------------------------------------"
    AppendLog " "
End If

If sRawCommand = "help" Then
    AppendLog "The help command requires a search string."
    AppendLog "EG:"
    AppendLog "    >help startservice"
End If

If sRawCommand = "resetconfig" Then
   Kill App.Path & "\xcyte.bin"
   DoEvents
   Shell App.Path & "\Restart.exe -noconfirm"
End If

If sRawCommand = "shutdownserver" Then
    ShutdownServer
End If
If Right(frmMain.SvrLog.Text, 1) = ">" Then
   ' Do something
   AppendLog "Bad command"
End If
If InStr(1, frmMain.SvrLog.Text, ">", vbTextCompare) <> 0 Then
   AppendLog "Bad or unrecognised console command!"
End If
AppendLog ">"
End Sub

Public Sub DoShutdownTimer(CountDown As String)
On Error Resume Next
If CountDown > 86400 Then
AppendLog "Server will shutdown in " & CountDown & " seconds, (" & Round(Int(CountDown / (86400)), 2) & " days)."
GoTo cse
End If

If CountDown > (60 * 60) Then
AppendLog "Server will shutdown in " & CountDown & " seconds, (" & Round(Int(CountDown / (60 * 60)), 2) & " hours)."
GoTo cse
End If
If CountDown > 60 Then
AppendLog "Server will shutdown in " & CountDown & " seconds, (" & Round(Int(CountDown / 60), 2) & " minutes)."
Else
AppendLog "Server will shutdown in " & CountDown & " seconds."
End If
cse:
AppendLog "Press F7 to cancel shutdown!"
AppendLog ">"
ShutdownCount = CountDown
frmMain.ShutdownTimer.Interval = 1000
frmMain.ShutdownTimer.Enabled = True
End Sub

Public Sub StopShutdown()
AppendLog "Shutdown terminated!"
AppendLog ">"
frmMain.ShutdownTimer.Enabled = False
End Sub
