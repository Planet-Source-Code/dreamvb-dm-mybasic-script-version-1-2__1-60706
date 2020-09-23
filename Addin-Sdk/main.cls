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
'Hello this is the add-ins template that you can use to make your own add-ins for DM MyBasic-Script
'Most of the template has already been setup up for you and is read to use.
'Just some little notes incase you have problums running any add-ins
'Each add-ins class must have it's propertie Name as main
'When your ready to compile your add-ins. the filename must be the same as
'you have the project name in Project->Properies
'For example this projects name is called devTest
'So when you compile you must compile it to devTest.dll

'Well that all you need to know for now.
'Ok installing the add-ins is easy all you need to do is drop the dll into
'the add-ins folder the when the IDE start your plug-ins will be shown in the add-ins toolbar

'If your deplaoying a add-ins to other people you can get them to installit or
' you can use a batch file and use the enviroment string to install for example.

'Copy yourAddin.dll %DM_MyBasic%\add-ins\yourAddin.dll

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
    ToolTipText = "TooltipText"
End Property

Public Sub InitAddin(mObject As Object)
    ' From here you can add your add-in code
    ' add forms, etc
    
    Set frmMain.Ide_Object = mObject
    frmMain.Show
End Sub
