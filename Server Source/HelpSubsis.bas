Attribute VB_Name = "HelpSubsis"
Option Explicit
Public Sub HandleHelpSubSystem(sHelpString As String)
If sHelpString = "startservice" Then
    AppendLog " "
    AppendLog "Syntax of startservice:"
    AppendLog "StartService [Servicename]"
    AppendLog " "
    AppendLog "Available services:"
    AppendLog "---------------------------------------------------"
    AppendLog "ActiveLog        - Active logging"
    AppendLog "ControlLog       - Active control logging"
    AppendLog "POP3             - Manages the pop3 protocol"
    AppendLog "SMTP             - Manages the smtp protocol"
    AppendLog "WEBMAIL          - Manages the webmail interface"
    AppendLog "---------------------------------------------------"
    AppendLog " "
    AppendLog "Note: To restart all services:"
    AppendLog "RestartService all"
End If

If sHelpString = "restartservice" Then
ListServices:
    AppendLog "Syntax of restartservice:"
    AppendLog "RestartService [Servicename]"
    AppendLog " "
    AppendLog "Available services:"
    AppendLog "---------------------------------------------------"
    AppendLog "ActiveLog        - Active logging"
    AppendLog "ControlLog       - Active data logging"
    AppendLog "POP3             - Manages the pop3 protocol"
    AppendLog "SMTP             - Manages the smtp protocol"
    AppendLog "WEBMAIL          - Manages the webmail interface"
    AppendLog "---------------------------------------------------"
    AppendLog " "
    AppendLog "Note: To restart all services:"
    AppendLog "RestartService all"
End If

If sHelpString = "shutdownserver" Then
    AppendLog "Syntax of shutdownserver:"
    AppendLog "ShutdownServer"
    AppendLog " "
    AppendLog "Warning: Executing this command will deny"
    AppendLog "all clients mail access until the application"
    AppendLog "is restarted!"
End If
End Sub
