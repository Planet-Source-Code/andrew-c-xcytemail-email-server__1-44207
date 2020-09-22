Attribute VB_Name = "Accounts"
Option Explicit
Dim cnn As ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim strConn As String

Public Function CreateAccount(sUsername As String, sPassword As String, saltEmail As String, ssmsemail As String) As Boolean
Dim errdesc As String
Dim strProvider As String
Dim strQuery As String
If AccountExists(sUsername) = True Then
   errdesc = "The " & sUsername & " mailbox already exists!" & vbCrLf & "Choose another name"
   GoTo ErrorOccured
End If
If Len(sUsername) < frmAdmin.uMinLength.Text Then
   errdesc = "The mailbox name is too short!" & vbCrLf & "The mailbox name must be longer than 3 characters."
   GoTo ErrorOccured
End If
If Len(sPassword) < frmAdmin.pMinLength.Text Then
   errdesc = "The password is too short!" & vbCrLf & "The password must be greater than 6 characters"
   GoTo ErrorOccured
End If

' One could use active directories to verify users, but i'm going to do it MY way :P

Set cnn = New ADODB.Connection
Set cmd = New ADODB.Command
Set rs = New ADODB.Recordset
strProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & GetWindir & "\xmbox\Accounts.mdb;Persist Security Info=False"
  
cnn.ConnectionString = strProvider
cnn.Open strProvider
cmd.ActiveConnection = cnn

rs.Open "Logins", cnn, adOpenDynamic, adLockOptimistic
rs.AddNew

rs!Username = sUsername
rs!Password = sPassword
rs!altEmail = saltEmail
rs!sms = ssmsemail
rs!Enabled = 1 'Enable the user account
rs.Update
rs.Close
cnn.Close

' Create the oldstyle login file
'Open subfolder(sUsername) & "\!Account.txt" For Output As #1
'Print #1, "1"
'Print #1, "1"
'Print #1, "1"

'Close #1

'CreateUserDbase sUsername
CreateAccount = True
Exit Function
ErrorOccured:
On Error Resume Next
CreateAccount = False
rs.Close
cnn.Close

End Function

Public Function AccountExists(sUsername As String) As Boolean
Dim strProvider As String
Dim strQuery As String
strProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & GetWindir & "\xmbox\Accounts.mdb;"
Dim objconn As ADODB.Connection
Dim objrs As ADODB.Recordset
Set objconn = New ADODB.Connection
Set objrs = New ADODB.Recordset

objconn.Open strProvider
strQuery = "SELECT * FROM Logins"
strQuery = strQuery & " WHERE Username = '" & sUsername & "'"
strQuery = strQuery & " ORDER BY " & "Username" & " ASC"
Set objrs = objconn.Execute(strQuery)
If objrs.EOF Then AccountExists = False: Exit Function
If objrs(0) <> vbNullString Then AccountExists = True
End Function

Public Function ValidatePassword(sUsername As String, sPassword As String) As Boolean
On Error GoTo Failed
Dim strProvider As String
Dim strQuery As String
strProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & GetWindir & "\xmbox\Accounts.mdb;"
Dim objconn As ADODB.Connection
Dim objrs As ADODB.Recordset
Set objconn = New ADODB.Connection
Set objrs = New ADODB.Recordset

objconn.Open strProvider
strQuery = "SELECT * FROM Logins"
strQuery = strQuery & " WHERE Username = '" & sUsername & "'"
strQuery = strQuery & " ORDER BY " & "Username" & " ASC"
Set objrs = objconn.Execute(strQuery)
If LCase(objrs(1)) = LCase(sPassword) Then
    ValidatePassword = True
Else
    ValidatePassword = False
End If
Exit Function
Failed:
    ValidatePassword = False
End Function

Public Sub CreateLoginsDatabase(Path As String)
'On Error GoTo ErrorCreateDB
'Stop
Dim Cat     As New ADOX.Catalog
Dim Tbl(6) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String

sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=""" & Path & """;"
Cat.Create sCnn

  '----------* Table Definition of Logins *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "Logins"
    .Columns.Append "Username", adVarWChar, 50
      .Columns("Username").Properties("Nullable").value = False
      .Columns("Username").Properties("Jet OLEDB:Allow Zero Length").value = True
    .Columns.Append "Password", adVarWChar, 50
      .Columns("Password").Properties("Nullable").value = False
      .Columns("Password").Properties("Jet OLEDB:Allow Zero Length").value = True
    .Columns.Append "Altemail", adVarWChar, 50
      .Columns("Altemail").Properties("Nullable").value = True
      .Columns("Altemail").Properties("Jet OLEDB:Allow Zero Length").value = True
    .Columns.Append "SMS", adVarWChar, 50
      .Columns("SMS").Properties("Nullable").value = True
      .Columns("SMS").Properties("Jet OLEDB:Allow Zero Length").value = True
    .Columns.Append "Enabled", adBoolean
      .Columns("Enabled").Properties("Nullable").value = False
      .Columns("Enabled").Properties("Jet OLEDB:Compressed UNICODE Strings").value = False
  End With
  '----------* Index Definitions of Logins *----------
  ReDim Idx(0)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "Username"
  Tbl(0).Indexes.Append Idx(0)

  Cat.Tables.Append Tbl(0)

  Set Cat = Nothing
  Exit Sub

ErrorCreateDB:
    msgErrR = MsgBox("    Error No. " & Err & " " & vbCrLf & Error, vbCritical + vbAbortRetryIgnore, "Database creation error!")
    Select Case msgErrR
      Case Is = vbAbort
      If Not (Cat Is Nothing) Then
        Set Cat = Nothing
      End If
      Exit Sub
     Case Is = vbRetry
       Resume Next
     Case Is = vbIgnore
       Resume
    End Select
End Sub


Public Sub CreateUserDbase(Username As String)
On Error GoTo ErrorCreateDB

Dim Cat     As New ADOX.Catalog
Dim Tbl(6) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String

sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=""" & GetWindir & "\xmbox\" & Username & ".mdb" & """;"
Cat.Create sCnn
  '----------* Table Definition of Logins *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "Emails"
    .Columns.Append "MailID", adVarWChar, 50
      .Columns("MailID").Properties("Nullable").value = False
      .Columns("MailID").Properties("Jet OLEDB:Allow Zero Length").value = True
    .Columns.Append "Timestamp", adVarWChar, 50
      .Columns("Timestamp").Properties("Nullable").value = False
      .Columns("Timestamp").Properties("Jet OLEDB:Allow Zero Length").value = True
    .Columns.Append "Subject", adVarWChar, 50
      .Columns("Subject").Properties("Nullable").value = True
      .Columns("Subject").Properties("Jet OLEDB:Allow Zero Length").value = True
    .Columns.Append "From", adVarWChar, 50
      .Columns("From").Properties("Nullable").value = True
      .Columns("From").Properties("Jet OLEDB:Allow Zero Length").value = True
  End With
  '----------* Index Definitions of Logins *----------
  ReDim Idx(0)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "MailID"
  Tbl(0).Indexes.Append Idx(0)

  Cat.Tables.Append Tbl(0)

  Set Cat = Nothing
  Exit Sub

ErrorCreateDB:
    msgErrR = MsgBox("    Error No. " & Err & " " & vbCrLf & Error, vbCritical + vbAbortRetryIgnore, "Database creation error!")
    Select Case msgErrR
      Case Is = vbAbort
      If Not (Cat Is Nothing) Then
        Set Cat = Nothing
      End If
      Exit Sub
     Case Is = vbRetry
       Resume Next
     Case Is = vbIgnore
       Resume
    End Select
End Sub

Public Sub AddEmail(Username As String, ID As String, Subject As String, Timestamp As String, from)
Dim strProvider As String
Dim strQuery As String
If fso.FileExists(GetWindir & "\xmbox\" & Username & ".mdb") = False Then
   CreateUserDbase Username
End If
Set cnn = New ADODB.Connection
Set cmd = New ADODB.Command
Set rs = New ADODB.Recordset
strProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & GetWindir & "\xmbox\" & Username & ".mdb;Persist Security Info=False"
  
cnn.ConnectionString = strProvider
cnn.Open strProvider
cmd.ActiveConnection = cnn

rs.Open "Emails", cnn, adOpenDynamic, adLockOptimistic
rs.AddNew
rs!MailID = ID
rs!Timestamp = Timestamp
rs!Subject = Subject
rs!from = from
rs.Update
rs.Close
cnn.Close
End Sub

Public Sub DeleteEmailIndex(Ref As String)

End Sub

Public Function GetEnabledState(sUsername As String) As Boolean
Dim strProvider As String
Dim strQuery As String
strProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & GetWindir & "\xmbox\Accounts.mdb;"
Dim objconn As ADODB.Connection
Dim objrs As ADODB.Recordset
Set objconn = New ADODB.Connection
Set objrs = New ADODB.Recordset

objconn.Open strProvider
strQuery = "SELECT * FROM Logins"
strQuery = strQuery & " WHERE Username = '" & sUsername & "'"
strQuery = strQuery & " ORDER BY " & "Username" & " ASC"
Set objrs = objconn.Execute(strQuery)
If (objrs(4)) = 0 Then
    GetEnabledState = True
Else
    GetEnabledState = False
End If
End Function
