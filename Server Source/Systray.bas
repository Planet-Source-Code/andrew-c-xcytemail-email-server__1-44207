Attribute VB_Name = "SystrayMod"
'//
'// Filename: SysTray.bas
'// Description: Contains functions for communicating
'//              with the system-tray
'//
Public Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer
 
Public Enum WM_CONSTANTS
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDBLCLK = &H203
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_RBUTTONDBLCLK = &H206
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
End Enum

Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Global Const NIM_ADD = 0
Global Const NIM_MODIFY = 1
Global Const NIM_DELETE = 2
Global Const NIF_MESSAGE = 1
Global Const NIF_ICON = 2
Global Const NIF_TIP = 4

Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, Flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
Dim nidTemp As NOTIFYICONDATA
nidTemp.cbSize = Len(nidTemp)
nidTemp.hwnd = hwnd
nidTemp.uID = ID
nidTemp.uFlags = Flags
nidTemp.uCallbackMessage = CallbackMessage
nidTemp.hIcon = Icon
nidTemp.szTip = Tip & Chr$(0)
setNOTIFYICONDATA = nidTemp
End Function

Public Sub AddIcon(ToolTip As String)
Dim i As Integer
Dim nid As NOTIFYICONDATA
nid = setNOTIFYICONDATA(hwnd:=frmMain.hwnd, ID:=vbNull, Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, CallbackMessage:=WM_MOUSEMOVE, Icon:=frmMain.Icon, Tip:=ToolTip)
i = Shell_NotifyIconA(NIM_ADD, nid)
End Sub

Public Sub ModIcon(ToolTip As String)
Dim i As Integer
Dim nid As NOTIFYICONDATA
nid = setNOTIFYICONDATA(hwnd:=frmMain.hwnd, ID:=vbNull, Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, CallbackMessage:=WM_MOUSEMOVE, Icon:=frmMain.Icon, Tip:=ToolTip)
i = Shell_NotifyIconA(NIM_MODIFY, nid)
End Sub

Public Sub RemoveIcon()
Dim i As Integer
Dim nid As NOTIFYICONDATA
nid = setNOTIFYICONDATA(hwnd:=frmMain.hwnd, ID:=vbNull, Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, CallbackMessage:=WM_MOUSEMOVE, Icon:=frmMain.Icon, Tip:="")
i = Shell_NotifyIconA(NIM_DELETE, nid)
End Sub

