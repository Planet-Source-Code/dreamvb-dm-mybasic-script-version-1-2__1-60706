VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function LoadPlug() As Integer
    'This just allows the IDE to know if the plug-ins was loaded
    LoadPlug = 1
End Function

Function iPic() As IPictureDisp
    'This is the picture function. that sends back an IPictureDisp picture
    ' This picture is then used for the add-ins button.
    ' Change your picture in the resource file.
    ' Note mask color for the bitmap is RGB(255,0,255)
    Set iPic = LoadResPicture(101, vbResBitmap)
End Function

Public Property Get ToolTipText() As String
    'Add the add-ins tooltip text below
    ' this text shows as you move over the add-in button
    ToolTipText = "Base Convertor"
End Property

Public Sub InitAddin(mObject As Object)
    ' From here you can add your add-in code
    ' add forms, etc
    
    Set frmConvert.Ide_Object = mObject
    frmConvert.Show
End Sub
