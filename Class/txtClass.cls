VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "txtClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private mTextBox As TextBox
Private mBackColor As OLE_COLOR
Private mForeColor As OLE_COLOR

Enum TEditMenu
    vsCUT
    vsDELETE
    vsCOPY
    vsPASTE
    vsSELALL
End Enum

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_SETMARGINS = &HD3
Private Const EC_LEFTMARGIN = &H1

Function GetColumn() As Long
    GetColumn = mTextBox.SelStart - SendMessage(mTextBox.hwnd, EM_LINEINDEX, GetCurrentLineNumber, ByVal 0&)
End Function

Function GetCurrentLineLength() As Long
    GetCurrentLineLength = SendMessage(mTextBox.hwnd, EM_LINELENGTH, _
    SendMessage(mTextBox.hwnd, EM_LINEINDEX, GetCurrentLineNumber, 0), 0)
End Function

Function GetCurrentLineNumber() As Long
    GetCurrentLineNumber = SendMessage(mTextBox.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) '+ 1
End Function

Public Function HighLightLine(LineNumber As Long)
Dim iLength As Long, lnIdx As Long
    
    GotoLine = LineNumber
    
    lnIdx = SendMessage(mTextBox.hwnd, EM_LINEINDEX, LineNumber, 0)
    iLength = SendMessage(mTextBox.hwnd, EM_LINELENGTH, lnIdx, 0)
    mTextBox.SelLength = iLength
    mTextBox.SetFocus
    
End Function


Public Function IsClipEmpty() As Boolean
    ' this just test to see if there any text on the clipboard
    IsClipEmpty = False
    If Trim(Len(Clipboard.GetText(vbCFText))) = 0 Then IsClipEmpty = True
End Function

Public Function EnableCutPaste() As Boolean
    If Len(Trim(mTextBox.SelText)) = 0 Then
        EnableCutPaste = False
        Exit Function
    Else
        EnableCutPaste = True
    End If
End Function

Public Function EditMenu(opCommands As TEditMenu)
    Select Case opCommands
        Case vsCUT
            Clipboard.SetText mTextBox.SelText
            mTextBox.SelText = ""
            mTextBox.SetFocus
        Case vsDELETE
            mTextBox.SelText = ""
            mTextBox.SetFocus
        Case vsCOPY
            Clipboard.SetText mTextBox.SelText
            mTextBox.SetFocus
        Case vsPASTE
            mTextBox.SelText = Clipboard.GetText
            mTextBox.SetFocus
        Case vsSELALL
            mTextBox.SelStart = 0
            mTextBox.SelLength = Len(mTextBox.Text)
            mTextBox.SetFocus
    End Select
End Function

Private Function SetMargin(nMarSize As Long)
    SendMessage mTextBox.hwnd, EM_SETMARGINS, EC_LEFTMARGIN, nMarSize
End Function

Public Property Let TextBox(ByVal vNewValue As Variant)
    Set mTextBox = vNewValue
End Property

Public Property Get Text() As String
    Text = mTextBox.Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    mTextBox.Text = vNewValue
End Property

Public Property Let GotoLine(vNewLineNum As Long)
Dim lnNum As Long
    'This will move the cursor to the value of vNewLineNum
    lnNum = SendMessage(mTextBox.hwnd, EM_LINEINDEX, ByVal vNewLineNum, ByVal 0&)
    If lnNum = -1 Then lnNum = 0
    mTextBox.SelStart = lnNum
    mTextBox.SetFocus
End Property

Public Property Get SelText() As String
    SelText = mTextBox.SelText
End Property

Public Property Let SelText(ByVal vNewValue As String)
    mTextBox.SelText = vNewValue
End Property

Public Property Get SelLength() As Long
    SelLength = mTextBox.SelLength
End Property

Public Property Let SelLength(ByVal vNewValue As Long)
    mTextBox.SelLength = vNewValue
End Property

Public Property Get SelStart() As Long
    SelStart = mTextBox.SelStart
End Property

Public Property Let SelStart(ByVal vNewValue As Long)
    mTextBox.SelStart = vNewValue
End Property

Public Sub SetFocus()
    mTextBox.SetFocus
End Sub

Public Property Let MarginSize(ByVal vNewValue As Long)
    SetMargin vNewValue
End Property

