VERSION 5.00
Begin VB.UserControl MX 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   510
   Begin VB.Label Title 
      Caption         =   "MX"
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   375
   End
End
Attribute VB_Name = "MX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Initialize()
    SetWinVersion
End Sub

Private Sub UserControl_Resize()
'    If UserControl.Width <> 32 Then
'        UserControl.Width = 400
'    End If
'    If UserControl.Height <> 32 Then
'        UserControl.Height = 250
'    End If
'    Title.Left = 0
End Sub

Public Function GetMX() As String
    If IsNetConnectOnline = True Then
        'If Not IsNetConnectViaProxy Then
            GetMX = MX_Query
            'AppendLog "MX lookup returned: " & GetMX
        'Else
        '    Err.Raise 0, "GetMX", "This computer is connected via a proxy server." & vbCrLf & "At this time, the wMX control does not support proxy servers."
        '    Exit Function
        'End If
    Else
    '    Err.Raise 0, "GetMX", "This computer is not currently connected to the internet."
        frmMain.Caption = "XCyteMail Network Server - Server is offline!"
        Debug.Print "Server is offline"
        'AppendLog "Server is not connected to internet."
        ' AppendLog ">"
        'If GetMX = "" Then GetMX = Domain
    End If
End Function

Public Property Get DNSCount() As Integer
    DNSCount = mi_DNSCount
End Property

Public Property Get MXCount() As Integer
    MXCount = mi_MXCount
End Property

Public Property Get PrefCount() As Integer
    PrefCount = mi_MXCount
End Property

Public Property Get Domain() As String
    Domain = ms_Domain
End Property

Public Property Let Domain(ByVal New_Domain As String)
    If Len(New_Domain) > 4 Then 'its a good host
        ms_Domain = New_Domain
    End If
End Property

Public Function DNS(ByVal Index As String) As String
    DNS = sDNS(Index)
End Function

Public Function MX(ByVal Index As String) As String
On Error Resume Next
    MX = sMX(Index)
End Function

Public Function Pref(ByVal Index As String) As String
    Pref = sPref(Index)
End Function


