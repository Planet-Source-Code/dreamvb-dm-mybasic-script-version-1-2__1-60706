VERSION 5.00
Begin VB.UserControl Bevel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ControlContainer=   -1  'True
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
End
Attribute VB_Name = "Bevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum eBevelStyle
    Lowered = 0
    Raised = 1
End Enum

Private Const def_BevelStyle As Integer = 1
Private m_BevelStyle As Integer

Private Sub DrawEffect(Direction As Integer)
    
    UserControl.Cls
    
    If Direction = 1 Then
        UserControl.Line (UserControl.ScaleWidth, 0)-(0, 0), vbWhite
        UserControl.Line (0, UserControl.ScaleHeight)-(0, -1), vbWhite
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), vbApplicationWorkspace
        UserControl.Line (UserControl.ScaleWidth, UserControl.ScaleHeight - 1)-(-1, UserControl.ScaleHeight - 1), vbApplicationWorkspace
    Else
        UserControl.Line (UserControl.ScaleWidth, 0)-(0, 0), vbApplicationWorkspace
        UserControl.Line (0, UserControl.ScaleHeight)-(0, -1), vbApplicationWorkspace
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), vbWhite
        UserControl.Line (UserControl.ScaleWidth, UserControl.ScaleHeight - 1)-(-1, UserControl.ScaleHeight - 1), vbWhite
    End If
    
    UserControl.Refresh
    
End Sub

Private Sub UserControl_Initialize()
    m_BevelStyle = def_BevelStyle
    DrawEffect def_BevelStyle
End Sub

Public Property Get BevelStyle() As eBevelStyle
    BevelStyle = m_BevelStyle
End Property

Public Property Let BevelStyle(ByVal vNewBevelStyle As eBevelStyle)
    m_BevelStyle = vNewBevelStyle
    DrawEffect m_BevelStyle
    PropertyChanged "BevelStyle"
End Property

Private Sub UserControl_Resize()
    DrawEffect m_BevelStyle
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BevelStyle = PropBag.ReadProperty("BevelStyle", def_BevelStyle)
End Sub

Private Sub UserControl_Show()
    DrawEffect m_BevelStyle
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BevelStyle", m_BevelStyle, def_BevelStyle)
End Sub

