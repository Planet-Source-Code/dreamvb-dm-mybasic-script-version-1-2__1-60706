Attribute VB_Name = "modMain"
Public DataPath As String
Public MainAppPath As String, Plg_Path As String
Public WebTemp As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public lzEngine_File As String, lzScript_File As String, lzBuild_Tool As String
Public ButtonPressed As Integer, TSelectionType As Integer, mGoto As Integer

Public IncFile As String
Public StartPage As String

Public Config_File As String
Public IniClass As New clsINI
Public Recent_File As String
Public RECENT_MAX_COUNT As Integer

Public clsTextBox As New txtClass
Public Const dlg_filter As String = "MyBasic Script(*.myb)|*.myb|Text Files(*.txt)|*.txt|"
Public Str_Error As String 'Error holder

Private Type Config
    nEngine As String
    nMergeTool As String
    nRecDocMax As Integer
    nFullSizeWindow As Integer
    nFont As String
    nFontStyle As Integer
    nFontSize As Integer
    nForeColor As Long
    nBackColor As Long
    nTabSize As Integer
    nLeftMargin As Integer
End Type

Public t_Config As Config

Function GetFileExt(lzFile As String) As String
    e_pos = InStrRev(lzFile, ".", Len(lzFile), vbBinaryCompare)
    If e_pos = 0 Then
        GetFileExt = lzFile
    Else
        GetFileExt = UCase(Mid(lzFile, e_pos + 1, Len(lzFile)))
    End If
End Function

Public Sub LoadFunctionsLst(tView As TreeView, lzListFile As String)
Dim sBuff As String, CatName As String, FuncName As String
Dim vLst As Variant, FuncLst As Variant
Dim sLine As String, e_pos As Integer, FoundSelection As Boolean
Dim Y As Boolean, x As Long, n As Integer, NodeCnt As Integer

On Error Resume Next

    'Check that the functions file is found
    If Not IsFileHere(lzListFile) Then
        MsgBox "File not found:" & vbCrLf & lzListFile, vbCritical, "File Not Found"
        Exit Sub
    End If
    
    tView.Nodes.Clear 'Clear treeview
    sBuff = OpenFile(lzListFile) 'Load the functions file
    
    vLst = Split(sBuff, vbCrLf) 'Split sBuff by vbcrlf
    sBuff = "" 'Clear up
    
    For x = 0 To UBound(vLst)
        sLine = LCase(Trim(vLst(x)))
        e_pos = InStr(1, sLine, "cat=", vbTextCompare)
        
        If e_pos > 0 Then
            CatName = Trim(Mid(vLst(x), e_pos + 4, Len(sLine))) 'Extract cat name
            NodeCnt = tView.Nodes.Count + 1 'Keep track of the treeview node count
            tView.Nodes.Add , tvwChild, "a:" & NodeCnt, CatName, 12, 12 'Add the cat
            FoundSelection = True 'Found a selection tag
            Y = False
        Else
            'Inside tag selection
            FoundSelection = False
            Y = True
        End If
        
        If (Y And FoundSelection = False) Then 'are we inside of a tag selection
        
        FuncLst = Split(vLst(x), ",") 'Split incomming line

        For n = 0 To UBound(FuncLst)
            FuncName = FuncLst(n) 'Get the function name
            If Len(FuncName) > 2 Then
                'Add the function name to the treeview
                tView.Nodes.Add NodeCnt, tvwChild, FuncName, FuncName, 11, 11
            End If
          Next n
        End If
    Next x
    
    'Clear up
    CatName = "": FuncName = "": sLine = ""
    Erase vLst: Erase FuncLst
    x = 0: e_pos = 0: n = 0: NodeCnt = 0

End Sub

Function GetFileTitle(lzFile As String) As String
Dim x As Integer, e_pos As Integer
    For x = 1 To Len(lzFile)
        If Mid(lzFile, x, 1) = "." Then e_pos = x
    Next x
    
    If e_pos = 0 Then
        GetFileTitle = lzFile
    Else
        GetFileTitle = Mid(lzFile, 1, e_pos)
    End If
    
End Function

Function GetFileName(lzPathFile As String) As String
Dim x As Integer, e_pos As Integer
    
    For x = 1 To Len(lzPathFile)
        If Mid(lzPathFile, x, 1) = "\" Then e_pos = x
    Next x
    
    x = 0
    
    If e_pos = 0 Then
        GetFileName = lzPathFile
    Else
        GetFileName = Mid(lzPathFile, e_pos + 1, Len(lzPathFile))
    End If
    
    e_pos = 0
    
    
End Function

Public Function GetAtom(AtomIdx As Integer) As String
Dim iBuff As String * 256
Dim iRet As Long

    iRet = GlobalGetAtomName(AtomIdx, iBuff, Len(iBuff))
    GetAtom = Left(iBuff, iRet)
    iBuff = "": iRet = 0
    
End Function

Function GetShPath(lpLongPath As String) As String
Dim iRet As Long
Dim sBuff As String * 256
    iRet = GetShortPathName(lpLongPath, sBuff, Len(sBuff))
    GetShPath = Left(sBuff, iRet)
    sBuff = ""
    
End Function

Function RunFile(lpFile As String, nHwnd As Long, lParm As String, WinOpenStyle As Integer)
    ShellExecute nHwnd, "open", lpFile, lParm, vbNullString, WinOpenStyle
End Function

Function FixPath(lzPath As String) As String
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function SaveFile(lzFile As String, FileData As String)
Dim iFile As Long
    iFile = FreeFile
    
    Open lzFile For Output As #iFile
        Print #iFile, FileData;
    Close #iFile

    FileData = ""
End Function

Public Function OpenFile(Filename As String) As String
Dim iFile As Long
Dim mByte() As Byte

    iFile = FreeFile
    Open Filename For Binary As #iFile
        ReDim Preserve mByte(0 To LOF(iFile))
        Get #iFile, , mByte
    Close #iFile
    
    OpenFile = StrConv(mByte, vbUnicode)
    
    Erase mByte
    
End Function

Function GetAbsPath(lpPath As String) As String
Dim x As Integer, e_pos As Integer
    For x = 1 To Len(lpPath)
        If Mid(lpPath, x, 1) = "\" Then e_pos = x
    Next
    
    If e_pos <> 0 Then GetAbsPath = Mid(lpPath, 1, e_pos)

End Function

Sub ConfigDefaults()
    'General selection
    IniClass.INIWriteKeyValue "General", "Engine", MainAppPath & "engine\engine.exe"
    IniClass.INIWriteKeyValue "General", "Merge", MainAppPath & "engine\mbComp.exe"
    IniClass.INIWriteKeyValue "General", "RecDocMax", 8
    IniClass.INIWriteKeyValue "General", "FullSizeWindow", 0
    'Editor selection
    IniClass.INIWriteKeyValue "Editor", "Font", "Courier"
    IniClass.INIWriteKeyValue "Editor", "FontSize", 10
    IniClass.INIWriteKeyValue "Editor", "FontStyle", 1
    IniClass.INIWriteKeyValue "Editor", "Fore-Color", 0
    IniClass.INIWriteKeyValue "Editor", "Back-Color", -2147483643
    IniClass.INIWriteKeyValue "Editor", "TabSize", 4
    IniClass.INIWriteKeyValue "Editor", "LeftMargin", 5
End Sub

Public Sub LoadConfig()
Dim sTmp As String
On Error Resume Next

    'General Selection
    sTmp = IniClass.INIReadKeyValue("General", "Engine", MainAppPath & "engine\engine.exe")
    sTmp = Replace(sTmp, "[APP_PATH]", MainAppPath)
    t_Config.nEngine = sTmp
    
    sTmp = IniClass.INIReadKeyValue("General", "Merge", MainAppPath & "engine\mbComp.exe")
    sTmp = Replace(sTmp, "[APP_PATH]", MainAppPath)
    t_Config.nMergeTool = sTmp
    t_Config.nFullSizeWindow = IniClass.INIReadKeyValue("General", "FullSizeWindow", 0)
    t_Config.nRecDocMax = IniClass.INIReadKeyValue("General", "RecDocMax", 8)
    
    t_Config.nBackColor = IniClass.INIReadKeyValue("Editor", "Back-Color", -2147483643)
    t_Config.nFont = IniClass.INIReadKeyValue("Editor", "Font", "Courier")
    t_Config.nFontStyle = IniClass.INIReadKeyValue("Editor", "FontStyle", 1)
    
    t_Config.nFontSize = IniClass.INIReadKeyValue("Editor", "FontSize", 10)
    t_Config.nForeColor = IniClass.INIReadKeyValue("Editor", "Fore-Color", 0)
    t_Config.nLeftMargin = IniClass.INIReadKeyValue("Editor", "LeftMargin", 5)
    t_Config.nTabSize = IniClass.INIReadKeyValue("Editor", "TabSize", 4)
    
    lzEngine_File = t_Config.nEngine
    lzBuild_Tool = t_Config.nMergeTool
    frmMain.txtCode.ForeColor = t_Config.nForeColor
    frmMain.txtCode.BackColor = t_Config.nBackColor
    frmMain.txtCode.Font = t_Config.nFont
    frmMain.txtCode.FontSize = t_Config.nFontSize
    
    frmMain.txtCode.FontBold = False
    frmMain.txtCode.FontItalic = False
    
    If t_Config.nFontStyle = 1 Then
        frmMain.txtCode.FontBold = True
    ElseIf t_Config.nFontStyle = 2 Then
        frmMain.txtCode.FontItalic = True
    ElseIf t_Config.nFontStyle = 3 Then
        frmMain.txtCode.FontItalic = True
        frmMain.txtCode.FontBold = True
    Else
        frmMain.txtCode.FontBold = False
        frmMain.txtCode.FontItalic = False
    End If
    
    RECENT_MAX_COUNT = t_Config.nRecDocMax - 1
    clsTextBox.MarginSize = t_Config.nLeftMargin
    
    sTmp = ""
End Sub

Public Sub WriteToConfig()
    'General selection
    IniClass.INIWriteKeyValue "General", "Engine", t_Config.nEngine
    IniClass.INIWriteKeyValue "General", "Merge", t_Config.nMergeTool
    IniClass.INIWriteKeyValue "General", "RecDocMax", Str(t_Config.nRecDocMax)
    IniClass.INIWriteKeyValue "General", "FullSizeWindow", Str(t_Config.nFullSizeWindow)
    'Editor selection
    IniClass.INIWriteKeyValue "Editor", "Font", t_Config.nFont
    IniClass.INIWriteKeyValue "Editor", "FontSize", Str(t_Config.nFontSize)
    IniClass.INIWriteKeyValue "Editor", "FontStyle", Str(t_Config.nFontStyle)
    IniClass.INIWriteKeyValue "Editor", "Fore-Color", Str(t_Config.nForeColor)
    IniClass.INIWriteKeyValue "Editor", "Back-Color", Str(t_Config.nBackColor)
    IniClass.INIWriteKeyValue "Editor", "TabSize", Str(t_Config.nTabSize)
    IniClass.INIWriteKeyValue "Editor", "LeftMargin", Str(t_Config.nLeftMargin)
End Sub
