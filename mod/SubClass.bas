Attribute VB_Name = "SubClass"
Public lpPrevWndProc As Long
Public mWnd As Long

Private Const COMPILE_RESULT As Long = &H512
Private Const GWL_WNDPROC = (-4)
Private Const WM_CLOSE = &H10

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Function GetConsolWnd() As Long
    'Get the current console window
    GetConsolWnd = FindWindow(vbNullString, "MyBasic-Script")
End Function

Public Sub CloseConsoleWnd(WndHangle As Long)
    'Send the close window message to the console
    SendMessage WndHangle, WM_CLOSE, ByVal 0&, ByVal 0&
End Sub

Public Sub Hook()
    'Place a hook on the window so we can catch the messages
    lpPrevWndProc = SetWindowLong(mWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHook()
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(mWnd, GWL_WNDPROC, lpPrevWndProc)
    'A must aways unhook a window. and Never hit the stop button in VB while hooked.
    ' unless VB will just crash simple.
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'This small function we use to catch any errors from the script engine
    ' The error message is then added to frmmain AddError were it's processed
    Select Case uMsg
        Case COMPILE_RESULT
            Str_Error = GetAtom(CInt(wParam))
            frmMain.AddError Str_Error
            Str_Error = ""
            Exit Function
        Case 120 'Good execute
            frmMain.Toolbar1.Buttons(12).Enabled = True ' Enable build button
            frmMain.Toolbar1.Buttons(11).Enabled = False 'Disable the stop button
        Case Else
            WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
            'above keep sending along the normal messages to the window
       End Select
   End Function
   
