Attribute VB_Name = "RunTimeCom"
Public cParams() As String
Public cRoot As String
Public cParCount As Integer

Public Sub GetParams(sString As String)
Dim sRawCommand As String
sRawCommand = sString & " "
cParCount = 0
' Seperate the command root and parameters
cRoot = Trim(LCase(Left(sRawCommand, InStr(1, sRawCommand, " ", vbTextCompare) - 1)))
sRawCommand = Right(sRawCommand, Len(sRawCommand) - Len(cRoot) - 1)

' Filter parameters into boxes
Do While InStr(1, sRawCommand, " ", vbTextCompare) <> 0
ReDim Preserve cParams(cParCount)
cParams(cParCount) = Left(sRawCommand, InStr(1, sRawCommand, " ", vbTextCompare) - 1)
sRawCommand = Right(sRawCommand, Len(sRawCommand) - Len(cParams(cParCount)) - 1)
If sRawCommand = " " Then sRawCommand = ""
cParCount = cParCount + 1
Loop

' Filter the last command into the box
ReDim Preserve cParams(cParCount)
cParams(cParCount) = sRawCommand
End Sub

Public Sub ProcessComLines()
Select Case LCase(cRoot)
Case "-adduser"
    Dim mName As String
    Dim mPword As String
    Dim altEmail As String
    Dim smsemail As String
    On Error GoTo BadStylax
    mName = cParams(0)
    mPword = cParams(1)
    On Error Resume Next
    altEmail = cParams(2)
    smsemail = cParams(3)
    CreateAccount mName, mPword, altEmail, smsemail
    End
BadStylax:
    ' Didn't work. Get the hell out!
    MsgBox "Must specify username and password." & vbCrLf & vbCrLf & "Syntax:" & vbCrLf & "Xcyte.exe -adduser Username Password"
    End
Case "-create"
    RuntimeStartup
    End
End Select
End Sub

Public Sub RuntimeStartup()
    'If fso.FolderExists(App.Path & "\email") = False Then MkDir (App.Path & "\email")
    If fso.FolderExists(App.Path & "\Config") = False Then MkDir (App.Path & "\Config")
    If fso.FolderExists(App.Path & "\Logs") = False Then MkDir (App.Path & "\Logs")
    Dim DBasePath As String
    DBasePath = GetWindir & "\xmbox"
    If fso.FolderExists(DBasePath) = False Then
       MkDir DBasePath
    End If
    
    Open subfolder("Filtered") & "\!account.txt" For Output As #1
    Print #1, "pw: Filtered"
    Print #1, "alt: postman@localhost"
    Print #1, "sms: postman@localhost"
    Close #1
    subfolder ("Out")
    Open App.Path & "\Config\Filters.cfg" For Append As #1
    Close #1
    Open App.Path & "\Config\Autoexec.txt" For Output As #1
    Print #1, "; Mail server autobooting script"
    Print #1, "StartService activelog"
    Print #1, "StartService pop3"
    Print #1, "StartService SMTP"
    Print #1, "StartService WEBMAIL"
    Print #1, "; Uncomment the following line to enable sidebar autohiding"
    Print #1, ";StartService SidebarAutoHide"
    Print #1, "; Uncomment the following line to make the app start in system tray"
    Print #1, ";StartService Systray"
    Print #1, "; Uncomment the following line to enable console logging"
    Print #1, "StartService keylog"
    Print #1, "message Welcome to your XCyteMail server!"
    Print #1, "message Server version: " & App.Major & "." & App.Minor & "." & App.Revision
    Print #1, "set ServerOwner=Mr. Smith"
    Print #1, "set ServerVersion=1.1.21"
    Print #1, "set ServerCode=Hybrid/X86XCyteMail"
    Close #1
    Open App.Path & "\Config\DNS.cfg" For Output As #1
    Close #1
    If fso.FolderExists(App.Path & "\config\Stubs") = False Then MkDir (App.Path & "\config\Stubs")
'    Stop
    ' Setup the databasing systems
    DBasePath = GetWindir & "\xmbox"
    If fso.FolderExists(DBasePath) = False Then
       MkDir DBasePath
    End If
    If fso.FileExists(DBasePath & "\Accounts.mdb") = False Then
    Accounts.CreateLoginsDatabase DBasePath & "\Accounts.mdb"
    End If
    MsgBox "Application setup completed!" & vbCrLf & "XCyteMail will now restart...", vbInformation, "Success"
    Shell "cmd.exe /c " & App.Path & "\Restart.exe -noappconfirm", vbNormalFocus
End Sub

