VERSION 5.00
Begin VB.UserControl Line3D 
   BackStyle       =   0  'Transparent
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   ScaleHeight     =   225
   ScaleWidth      =   5085
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4950
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   4950
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "Line3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Initialize()
resizeLine
End Sub

Private Sub resizeLine()
' Resize the line3D control
Line1.X1 = 0
Line2.X1 = 0
Line1.X2 = UserControl.Width
Line2.X2 = UserControl.Width

End Sub

Private Sub UserControl_Resize()
Call resizeLine
End Sub
