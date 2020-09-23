Attribute VB_Name = "modComp"
Function GetFileTitle(lzFile As String) As String
Dim x As Integer, e_pos As Integer
    'Extract a filename all the way to the dot
    'eg GetFileTitle(C:\folder\this\go.txt) returns C:\folder\this\go.
    
    For x = 1 To Len(lzFile)
        If Mid(lzFile, x, 1) = "." Then e_pos = x
    Next
    
    If e_pos = 0 Then
        GetFileTitle = lzFile
    Else
        GetFileTitle = Mid(lzFile, 1, e_pos)
    End If
    
End Function

Function GetAbsPath(lpPath As String) As String
Dim x As Integer, e_pos As Integer
    For x = 1 To Len(lpPath)
        If Mid(lpPath, x, 1) = "\" Then e_pos = x
    Next
    
    If e_pos <> 0 Then GetAbsPath = Mid(lpPath, 1, e_pos)

End Function

Function FixPath(lzPath As String) As String
    'Fix a pth by adding a backslash
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    'Check if a given file is found
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Public Function OpenFile(Filename As String) As String
Dim iFile As Long
Dim mByte() As Byte
    'Open a file for reading
    iFile = FreeFile
    Open Filename For Binary As #iFile
        ReDim Preserve mByte(0 To LOF(iFile))
        Get #iFile, , mByte
    Close #iFile
    
    OpenFile = StrConv(mByte, vbUnicode)
    
    Erase mByte
    
End Function

Sub Main()
Dim s_Cmd As String, StrA As String, StrEndBuff As String, sLen1 As String, nStr As String
Dim Script_Engine As String, OutFile As String, TFile As Long

    s_Cmd = Trim(Command) 'Get command line argv
    
    If Len(s_Cmd) = 0 Then
        'Check that the command lines is not empty
        MsgBox "Incorrect command line arguments" _
        & vbCrLf & "Use infile.myb - outfile.exe", vbInformation, "Error"
        End
    End If
    
    s_Cmd = Replace(s_Cmd, Chr(34), "", , , vbBinaryCompare) 'Get rid of the char 34
        
    If FileLen(s_Cmd) = 0 Then
        'Empty file not much point with going on
        MsgBox "Script file is empty unable to compile.", vbCritical, "Error"
        s_Cmd = ""
        End
    End If
    
    StrA = OpenFile(s_Cmd) 'Open the script file
    sLen1 = Len(StrA) 'Get the length of the script file
    nStr = sLen1 & Space((12 - Len(sLen1))) 'File Size and extra spaces used for header info
    StrEndBuff = StrA & nStr & Chr(25) 'Final data
    
    Script_Engine = FixPath(App.Path) & "engine.exe" 'Link to my Basic-Engine
    
    If Not IsFileHere(Script_Engine) Then
        'File was not found so end
        MsgBox "Engine.exe not found.", vbCritical, "File Not Found"
        GoTo Clean:
    End If
    
    'Make a copy of the engine to the same location as the script file
    'also rename the engine.exe to the script filename.exe
    
    OutFile = GetFileTitle(s_Cmd) & "exe" 'The path and name of the exe to write
    FileCopy Script_Engine, OutFile 'Do the file copy
    
    'Write the script to the new copyed file
    TFile = FreeFile 'Free file pointer
    'Dump script and header info to the newly copyed file
    Open OutFile For Binary As #TFile
        Put #TFile, LOF(TFile), StrEndBuff
    Close #TFile
    'Clear up and end
    
Clean:
    Script_Engine = ""
    StrEndBuff = ""
    OutFile = ""
    nStr = ""
    StrA = ""
    sLen1 = ""
    s_Cmd = ""
    End
    
End Sub
