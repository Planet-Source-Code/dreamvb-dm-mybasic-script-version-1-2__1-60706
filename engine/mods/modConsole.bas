Attribute VB_Name = "modConsole"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal dwCursorPosition As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function FillConsoleOutputCharacter Lib "kernel32.dll" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As Long, lpNumberOfCharsWritten As Long) As Long
Private Declare Function FillConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As Long, lpNumberOfAttrsWritten As Long) As Long

Private Type COORD
    x As Integer
    y As Integer
End Type

Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Enum TextAttr
    'forecolors
    FOREGROUND_RED = &H4
    FOREGROUND_GREEN = &H2
    FOREGROUND_BLUE = &H1
    FOREGROUND_INTENSITY = &H8
    'backcolors
    BACKGROUND_RED = &H40
    BACKGROUND_GREEN = &H20
    BACKGROUND_BLUE = &H10
    BACKGROUND_INTENSITY = &H80
End Enum

Private hInput As Long
Private hOutput As Long

Private cBuff As CONSOLE_SCREEN_BUFFER_INFO

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&

Public Function InitConsole() As Boolean
    InitConsole = False

    If AllocConsole Then
        hInput = GetStdHandle(STD_INPUT_HANDLE)
        hOutput = GetStdHandle(STD_OUTPUT_HANDLE)
        InitConsole = (hInput And hOutput)
    End If
    
End Function

Public Sub cCls(ResetCurPos As Boolean)
Dim m_Width As Long, m_Height As Long

    If GetConsoleBuffInfo Then
        m_Width = (cBuff.srWindow.Right - cBuff.srWindow.Left + 1)
        m_Height = (cBuff.srWindow.Bottom - cBuff.srWindow.Top + 1)
        FillConsoleOutputCharacter hOutput, 32, m_Width * m_Height, ByVal 0, 0
        FillConsoleOutputAttribute hOutput, cBuff.wAttributes, m_Width * m_Height, ByVal 0, 0
        
        If ResetCurPos Then cSetCursorPosition 0, 0
        
    End If
    
    m_Width = 0: m_Height = 0
    
End Sub

Public Sub cSetCursorPosition(x As Integer, y As Integer)
Dim mCoord As Long
Dim mCoordPrt As Long
    'Sets the position of the cursor on the console
    mCoordPrt& = VarPtr(mCoord) 'Get the address of mCoord
    CopyMemory mCoord, x, 2
    CopyMemory ByVal mCoordPrt& + 2, y, 2
    SetConsoleCursorPosition hOutput, mCoord
End Sub

Private Function GetConsoleBuffInfo() As Boolean
    ' Get Console screen buffer info
    GetConsoleBuffInfo = False
    GetConsoleBuffInfo = GetConsoleScreenBufferInfo(hOutput, cBuff) <> 0
End Function

Public Function cReadConsole() As String
On Error Resume Next
    Dim lpText As String * 256
    
    If ReadConsole(hInput, lpText, Len(lpText), vbNull, vbNull) <> 0 Then
        cReadConsole = Left(lpText, InStr(1, lpText, Chr(0)) - 3)
    End If
    
End Function

Public Sub cBeep()
    cWrite Chr(7)
End Sub

Public Sub cSetTitle(lpTitle As String)
    SetConsoleTitle lpTitle
End Sub

Public Sub cWrite(Optional lpText As String = "")
    WriteConsole hOutput, lpText, Len(lpText), ByVal 0, ByVal 0
End Sub

Public Sub cWriteLine(Optional lpText As String = "")
    cWrite lpText & vbCrLf
End Sub

Public Sub cPause()
    Call cWriteLine
    Call cWriteLine("Press any key to continue . . .")
    Call cReadConsole
End Sub

Public Function cFree()
    CloseConsole = FreeConsole
End Function

Public Property Let TextAttribute(ByVal vNewTxtAttr As TextAttr)
    SetConsoleTextAttribute hOutput, vNewTxtAttr
End Property
