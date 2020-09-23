VERSION 5.00
Begin VB.UserControl Sep 
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ScaleHeight     =   3495
   ScaleWidth      =   4815
   ToolboxBitmap   =   "Sep.ctx":0000
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   5564
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   15
      X2              =   5564
      Y1              =   15
      Y2              =   15
   End
End
Attribute VB_Name = "Sep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'©Copyright HardStream Software
'www.hardstream.tk
'Info@HardStream.tk
'
'This is a usercontrol that I wrote myself.
'I know it isn't that great but it can be useful sometimes
'This is a nice code for beginners...

Property Get TopLineColor() As OLE_COLOR
TopLineColor = Line1.BorderColor
End Property

Property Let TopLineColor(Clr As OLE_COLOR)
Line1.BorderColor = Clr
PropertyChanged TopLineColor
End Property

Property Get BottomLineColor() As OLE_COLOR
BottomLineColor = Line2.BorderColor
End Property

Property Let BottomLineColor(Clr As OLE_COLOR)
Line2.BorderColor = Clr
PropertyChanged BottomLineColor
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Line1.BorderColor = PropBag.ReadProperty(TopLineColor, &H808080)
Line2.BorderColor = PropBag.ReadProperty(BottomLineColor, vbWhite)
End Sub

Private Sub UserControl_Resize()
Line1.X2 = UserControl.Width
Line2.X2 = UserControl.Width
UserControl.Height = Line2.Y1 + 10
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty TopLineColor, Line1.BorderColor, &H808080
PropBag.WriteProperty BottomLineColor, Line2.BorderColor, vbWhite
End Sub
