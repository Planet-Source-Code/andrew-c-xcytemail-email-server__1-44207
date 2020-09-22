Attribute VB_Name = "MDBFunctions"
Option Explicit
Public Sub CreateEmptyDatabase(Path As String)
On Error GoTo ErrorCreateDB
Dim Cat     As New ADOX.Catalog
Dim Tbl As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Jet OLEDB:Engine Type=5;Data Source=""" & Path & """"
Cat.Create sCnn
Exit Sub
ErrorCreateDB:
MsgBox "Failed in creating database file! " & vbCrLf & "Reason: " & vbCrLf & Err.Description, vbInformation, "JET Database Engine"
Exit Sub
End Sub

Public Sub CreateTable(Path As String, TableName As String, FieldString As String)
Dim strProvider As String
Dim strQuery As String
On Error GoTo ctError
strProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & Path & ";"
Dim objconn As ADODB.Connection
Dim objrs As ADODB.Recordset
Set objconn = New ADODB.Connection
Set objrs = New ADODB.Recordset

objconn.Open strProvider
strQuery = "CREATE TABLE " & TableName & " " & FieldString
Set objrs = objconn.Execute(strQuery)
Exit Sub
ctError:
MsgBox "Failed!" & vbCrLf & "Reason: " & Err.Description
End Sub
