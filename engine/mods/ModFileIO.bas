Attribute VB_Name = "ModFileIO"
'This bas file is for dealing with the FileIO Operations

Enum FileIOOptions 'File operation types
    m_Kill = 0
    m_mkDir = 1
    m_FileCpy = 2
    m_FileRn = 3
    m_chDir = 4
    m_RmDir = 5
    m_Dir = 6
    mShell = 7
    mSendKey = 8
    mInput = 9
End Enum

Function FileIOOperation(FileIoOption As FileIOOptions, Optional lpValue As Variant) As Variant
On Error GoTo FileIOErr:
Dim Temp As Variant
Dim ParmList(1) As String 'Used to hold the parms of the function

    'This function deals with different file operations
    
    Select Case FileIoOption
        Case m_Kill 'Kill filename
            If isEmptyLine(ProcessLine) Then Abort 11, CurrentLine, "Kill", "Kill srcfile"
            Kill Eval(ProcessLine)
        Case m_mkDir 'Make folder
            If isEmptyLine(ProcessLine) Then Abort 11, CurrentLine, "MkDir", "MkDir FolderName"
            MkDir Eval(ProcessLine) 'Kill the file
        Case m_chDir
            If isEmptyLine(ProcessLine) Then Abort 11, CurrentLine, "ChDir", "ChDir FolderPath"
            ChDir Eval(ProcessLine) 'Change Dir path
            SetConst ConstIndex("app.path"), Eval(ProcessLine) 'Assign app.path const new path change
        Case m_RmDir 'Remove folder
            If isEmptyLine(ProcessLine) Then Abort 11, CurrentLine, "RmDir", "RmDir FolderName"
            RmDir Eval(ProcessLine)
        Case m_Dir
            GetTwoParms CStr(lpValue), ParmList, True
            If ParmList(1) = vbNullChar Then
                FileIOOperation = Dir(ParmList(0))
            Else
                FileIOOperation = Dir(ParmList(0), ParmList(1))
            End If
        Case mShell 'Shell to open a file
            If isEmptyLine(ProcessLine) Then Abort 11, CurrentLine, "Shell", "Shell PathName, [WindowStyle]"
            
            GetTwoParms CStr(lpValue), ParmList, True
            If ParmList(1) = vbNullChar Then
              FileIOOperation = Shell(ParmList(0))
            Else
                FileIOOperation = Shell(ParmList(0), ParmList(1))
            End If
        Case mSendKey 'send keys
            If isEmptyLine(ProcessLine) Then Abort 11, CurrentLine, "Shell", "Shell PathName, [WindowStyle]"
            GetTwoParms CStr(lpValue), ParmList, True
            If ParmList(1) = vbNullChar Then SendKeys ParmList(0) Else SendKeys ParmList(0), ParmList(1)
        Case mInput
            GetTwoParms CStr(lpValue), ParmList
            Temp = Trim(ParmList(1))
            
            If Left(Temp, 1) = "#" Then
                Temp = Eval(Right(Temp, Len(Temp) - 1))
            End If
            
            FileIOOperation = Input$(ParmList(0), Temp)
            
    End Select
    
    Erase ParmList
    Exit Function
FileIOErr:
    'This will soon be removed to allow custom error checking from a script
    If Err Then Abort 2, CurrentLine, Err.Description

End Function

Sub DoFileCopy()
Dim ParmList(1) As String
On Error GoTo FileIOErr:

    'Sub used to preform a filecopy
    If isEmptyLine(ProcessLine) Then Abort 11, CurrentLine, "FileCopy", "FileCopy srcfile, destfile"
    GetTwoParms ProcessLine, ParmList, False 'Get the Parms
    FileCopy ParmList(0), ParmList(1) 'Do the filecopy
    Erase ParmList 'Erase parm list
    Exit Sub
    
FileIOErr:
    'This will soon be removed to allow custom error checking from a script
    If Err Then Abort 2, CurrentLine, Err.Description

End Sub

Sub DoRename()
Dim e_pos As Integer
Dim A As String, B As String
On Error GoTo FileIOErr:
    'Sub used to rename a file
    
    If isEmptyLine(ProcessLine) Then Abort 9, CurrentLine, vbCrLf & "USE: OldName As NewName"
    e_pos = StrFind(0, ProcessLine, "as") 'Find the as keyword
    If e_pos = 0 Then Abort 10, CurrentLine 'Exit is above is not met
 
    A = Eval(Trim(Left(ProcessLine, e_pos - 1))) 'Extract first parm oldfilename
    B = Eval(Trim(Mid(ProcessLine, e_pos + 2, Len(ProcessLine) - e_pos - 1))) 'Extract seconed part newfilename
    Name A As B 'Rename the file
    
    'Clear up
    A = "": B = "": e_pos = 0
    Exit Sub
FileIOErr:
    'This will soon be removed to allow custom error checking from a script
    If Err Then Abort 2, CurrentLine, Err.Description
End Sub

Public Sub DoFileInput()
Dim e_pos As Integer, lpVar_Idx As Integer, lVarName As String
Dim inData As String
' On Error GoTo FileIOErr:
    
    'This sub is used for Inputting data from a file
    
    e_pos = CharPos(ProcessLine, ",") 'Find parm seperator
    If e_pos = 0 Then Abort 2, CurrentLine, "Required: #"
    FilePointer = Eval(Trim(Left(ProcessLine, e_pos - 1)))
    
    lVarName = Trim(Mid(ProcessLine, e_pos + 1, Len(ProcessLine))) 'Extract variable name
    lpVar_Idx = VariableIndex(lVarName) 'Get the above variables index
    If lpVar_Idx = -1 Then Abort 6, CurrentLine, lVarName 'No Exit so we exit
    
    Input #FilePointer, inData 'Do the File Input
    SetVariableData lVarName, inData 'Assign input to the variable
    
    'Clean up
    lVarName = ""
    inData = ""
    e_pos = 0
    Exit Sub
    
'FileIOErr:
    'This will soon be removed to allow custom error checking from a script
  '  If Err Then Abort 2, CurrentLine, Err.Description
    
End Sub

Private Sub DoFilePrint(FileData As Variant, Optional lpParm As Variant)
On Error GoTo FileIOErr:

    'This Sub is used to do a file print
    Select Case FileMode
        Case "OUTPUT" 'Print the data to the file
            If lpParm = vbNullString Then
                Print #FilePointer, FileData
            Else
                Print #FilePointer, lpParm, FileData
            End If
        Case "APPEND" 'Append data to exsiting data in a file
            If lpParm = vbNullString Then
                Print #FilePointer, FileData
            Else
                Print #FilePointer, lpParm, FileData
            End If
    End Select
    
    Exit Sub
FileIOErr:
    'This will soon be removed to allow custom error checking from a script
    If Err Then Abort 2, CurrentLine, Err.Description
End Sub

Public Sub GetIOMethod()
Dim iParmCnt As Integer, sLine As String, vStr As Variant

    'Sub used to identify File Mode, File Data and File Pointer
    iParmCnt = CountIF(ProcessLine, ",") 'Find parm seperator
    
    If iParmCnt = -1 Then Abort 2, CurrentLine, "Syntax Error" 'Exit if no parms are found
    
    If iParmCnt = 0 Then
        'Extract file Pointer
       vStr = Split(ProcessLine, ",")
       FilePointer = Eval(vStr(0)) 'File Pointer
       vStr(1) = Eval(vStr(1)) 'FileData
       DoFilePrint vStr(1), vbNullString
    Else
        vStr = Split(ProcessLine, ",")
        FilePointer = vStr(0) 'File Pointer
        vStr(1) = Eval(vStr(1)) 'other
        vStr(2) = Eval(vStr(2)) 'Filedata
        DoFilePrint vStr(1), vStr(2)
    End If
    
    'Clear up
    sLine = ""
    iParmCnt = -1
    Erase vStr
    
End Sub

Sub SetupFileMode(lzFile As String)
On Error GoTo FileIOErr:

    'Open a file in the selected users file mode
    Select Case FileMode
        Case "INPUT"
            Open lzFile For Input As FilePointer
        Case "OUTPUT"
            Open lzFile For Output As FilePointer
        Case "APPEND"
            Open lzFile For Append As FilePointer
    End Select

    Exit Sub
    
FileIOErr:
    'This will soon be removed to allow custom error checking from a script
    If Err Then Abort 2, CurrentLine, Err.Description
End Sub

Sub DoFileIO()
Dim e_pos As Integer, d_pos As Integer
Dim sTemp As Variant, lzFile As String
Dim vHolder As Variant, vStrHold() As String, iSize As Integer

    ReDim vStrHold(0)
    iSize = -1
    
    If isEmptyLine(ProcessLine) Then Abort 9, CurrentLine

    e_pos = CharPos(ProcessLine, "[") 'Check for opening bracket
    d_pos = CharPos(ProcessLine, "]") 'Check for closeing bracket
    
    If Not (e_pos > 0) Or Not (d_pos > 0) Then Abort 9, CurrentLine 'No opening and closeing brackets were found
    'Extract the file name using brakcet positions from above
    lzFile = Eval(Trim(Mid(ProcessLine, e_pos + 1, d_pos - e_pos - 1)))
    
    sTemp = UCase(Trim(Mid(ProcessLine, d_pos + 1, Len(ProcessLine) - d_pos)))
    
    vHolder = Split(sTemp, Chr$(32), , vbBinaryCompare)
    
    For d_pos = 0 To UBound(vHolder)
        sLine = Trim(vHolder(d_pos))
        If Len(sLine) > 0 Then
            iSize = iSize + 1
             ReDim Preserve vStrHold(iSize)
             vStrHold(iSize) = sLine
        End If
    Next
    
    Erase vHolder
    sLine = ""
    
    iSize = UBound(vStrHold)
    If iSize <> 3 Then Abort 10, CurrentLine
    If Not vStrHold(0) = "FOR" Then Abort 10, CurrentLine 'Check for FOR Keyword
    
    If Not (vStrHold(1) = "INPUT" Or vStrHold(1) = "OUTPUT" Or vStrHold(1) = "APPEND") Then
        Abort 10, CurrentLine  'Check for vaild file mode
    End If
    
    If Not vStrHold(2) = "AS" Then Abort 10, CurrentLine  'Check for AS Keyword
    If Len(vStrHold(3)) = 0 Then Abort 10, CurrentLine  'Check for file pointer
    
    sTemp = vStrHold(3)
    If Left(sTemp, 1) = "#" Then sTemp = RevStrLeft(CStr(sTemp), 1)
    FileMode = vStrHold(1)
    FilePointer = Eval(sTemp)
    'Lets carry open
    SetupFileMode lzFile
    
    'Clear up
    Erase vStrHold()
    e_pos = 0: d_pos = 0: iSize = 0
    lzFile = "": sTemp = ""
    
End Sub
