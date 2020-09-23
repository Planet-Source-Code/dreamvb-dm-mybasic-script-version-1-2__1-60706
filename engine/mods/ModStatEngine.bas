Attribute VB_Name = "ModStatEngine"
' This is the main mod for the start up events for the engine.

Private Sub Execute1(lzCommand As String)
Dim sCode As String
    If Len(lzCommand) = 0 Then
        'No command argv were found so we inform the user and exit
        MsgBox "Incorrect command line arguments" _
        & vbCrLf & "USE: " & App.EXEName & ".exe" & " ProgramName.myb", vbCritical, "Error"
        End
    Else
        If Left(lzCommand, 1) = Chr(34) And Right(lzCommand, 1) = Chr(34) Then
            'Fix the filename by removeing any quotes
            lzCommand = Right(lzCommand, Len(lzCommand) - 1): lzCommand = Left(lzCommand, Len(lzCommand) - 1)
        End If
        'Check the that the file does exsist
        If Not IsFileHere(lzCommand) Then
            'No file was found so we just exit
            MsgBox "File not found:" & vbCrLf & lzCommand, vbCritical, "Error"
            lzCommand = ""
            End
        Else
            'load in the script file
            sCode = OpenFile(lzCommand)
            'Now that we have the code we can now run the script
            AddCode sCode
            lzCommand = "" 'Clear the command line
        End If
    End If
End Sub

Public Sub RunCode(lpScript As String)
    GetLineCount lpScript 'Get the number of lines from the script
    
    If LineCount <= 0 Then
        'Abort if no lines were found in the script
        Abort 0, 0
        Exit Sub
    Else
        GetCodeLines lpScript 'Load in the code lines
        InitConsole 'Start the console
        cSetTitle "MyBasic-Script" 'Set the console title
        pngParser 'Start the executeing of the code
    End If
    
End Sub

Public Sub pngParser()
    'What this sub does is loop though all the code lines in LineHolder
    ' We then call TokenKeywords that returns keywords from the current
    ' executeing line eg PRINT,CLS,BEEP etc
    ' we then call Parser this then executes the current keywords
    
    Do While CurrentLine < LineCount 'Loop until we hit the max number of lines in the script
        CurrentLine = CurrentLine + 1 'Keep a count on our CurrentLine
        TokenKeywords 'call TokenKeywords
        Parser 'Call Parser
        DoEvents ' Allow other things to process
    Loop
    
End Sub

Public Sub Parser()
Dim thisToken As String
    thisToken = ""
    Static inc As Integer
    
    thisToken = TokenKeywords 'Get the current token or the current line been processed

    If Len(thisToken) <> 0 Then
        Select Case thisToken
            Case "CLOSE"
                If FilePointer <> -1 Then Close #FilePointer: FilePointer = -1
            Case "REM"
                thisToken = "" 'Comments here we do nothing
            Case "DO"
                Call DoLoop
            Case "DOEVENTS"
                 DoEvents 'Call do
            Case "DIM"
                DoDim ProcessLine
            Case "BEEP"
                cBeep 'Beep
            Case "CASE"
                Call DoCase 'Case
            Case "CHDIR"
                FileIOOperation m_chDir 'Change path
            Case "CONST"
                Call DoConst 'Consts
            Case "COLOR"
                Call doColor
            Case "CLEAR"
                Call DoClearVars 'Clear Variables
            Case "CLS"
                cCls True 'Clears all the console
            Case "CLG"
                cCls False 'Clear the screen up to last cursor postion
            Case "FOR"
                Call DoFor
            Case "FILECOPY"
                Call DoFileCopy
            Case "GOTO"
                Call doGoto 'Goto
            Case "INPUT"
                Call DoInput
            Case "KILL"
                FileIOOperation m_Kill 'Kill file
            Case "LET"
                DoAssign1 ProcessLine, True 'Assignment Let
            Case "LOCATE"
                Call Locate 'Locate used to position the text in the console
            Case "OPEN"
                Call DoFileIO
            Case "PRINT"
                Call PrintA ' Print Statement
            Case "RMDIR"
                FileIOOperation m_RmDir 'Remove Folder
            Case "SELECT"
               Call DoSelect 'Start of Select case
            Case "MKDIR"
                FileIOOperation m_mkDir  'make folder
            Case "NAME"
                Call DoRename
            Case "PAUSE"
                Call cPause ' pause the console
            Case "SHELL"
                FileIOOperation mShell, ProcessLine
            Case "SENDKEYS"
                FileIOOperation mSendKey, ProcessLine
            Case "LOOP"
                Call DoEndOfLoop 'End of Loop
            Case "EXIT"
                Call DoExit
            Case "END"
                'End Program and clean up
                cFree ' free the console
                Reset ' reset
                thisToken = "" ' clear current token
            Case "NEXT"
                Call DoForNext 'Call DoForNext
            Case Else
                'we have an unkown command but wait it may be a variable lets see
                If CharPos(ProcessLine, "=") <> 0 And VariableIndex(thisToken) = -1 Then
                    'if the string has an assign and is not a variable
                    ' we must infoprm the user of the error
                    'Check is the assignment is a const if it is. We must show the error message
                    If ConstIndex(thisToken) <> -1 Then Abort 2, CurrentLine, "Assignment to constant not permitted"
                    Abort 6, CurrentLine, thisToken
                    Exit Sub
                End If
                
                If VariableIndex(thisToken) = -1 Then
                    'If assign was found and the no varaible was found we inform of the error
                    If Right(thisToken, 1) = ":" = True Then Exit Sub
                    If (ProcessLine = vbNullChar) Then ProcessLine = ""
                    Abort 1, CurrentLine, thisToken & " " & ProcessLine
                    Exit Sub
                Else
                    'Ok we have an assign and a vaild variable so lets carry on
                    DoAssign1 ProcessLine, False, thisToken
                    Exit Sub
                End If
        End Select
    End If
    
End Sub

Public Function TokenKeywords() As String
Dim x_pos As Integer, h_pos As Integer, sLine As String

    ' this function is used to process the current line and find any tokens
    ' the function works by looking for a white space in the current lines
    ' ex
    ' <KeyWord> |Space| <KeyWord data> ex PRINT "HELLO WORLD"
    
    sLine = Trim(LineHolder(CurrentLine)) 'Trim down the line
    x_pos = InStr(1, sLine, Chr$(32), vbBinaryCompare) 'Locate the space chr(32)
    
    If x_pos > 0 Then ' Yes we have we found a space
        TokenKeywords = UCase(Mid(sLine, 1, x_pos - 1)) 'Get and return the keyword
        ProcessLine = Trim(Mid(sLine, x_pos + 1, Len(sLine))) ' Get the keywords data
    Else ' OK no space was found so we must asume this is a keyword with no data eg BEEP,CLS etc
        If Len(sLine) <> 0 Then
            ProcessLine = vbNullChar 'Clear the process line as we have no data for this keyword
            TokenKeywords = UCase(sLine) ' Get and return the Keyword
        End If
    End If
    
    'Clean up used vars
    x_pos = 0
    sLine = ""
    
End Function

Public Sub AddCode(lpCodeScript As String)
    Reset 'Call Global Reset
    AddSystemVars 'Add the system consts
    RunCode lpCodeScript 'Add the script code to be run
    cPause 'Put a pause on the console so it does not flash of
    cFree  'Clsoe the console we script has finsihed
    
    'The code below is used to send a 120 good message to the IDE
    ' to inform it that the script has finsihed
    If GetIde <> 0 Then
        SendMessage GetIde, 120, ByVal 0, ByVal 0
    End If
    
    End ' close the engine
End Sub

Sub Main()
Dim s_Cmd As String, ThisFile As String
Dim iFile As Long, sHead As String, iSize As Long
Dim Buffer As String

    ThisFile = FixPath(App.Path) & App.EXEName & ".exe"
    iFile = FreeFile
    s_Cmd = Trim(Command$) 'Get the command lines argv
    
    sHead = String(13, Chr(0)) 'Create a buffer
    'Open this file in binary mode
    Open ThisFile For Binary As #iFile
        Get #iFile, LOF(iFile) - 12, sHead 'Extract 12 bytes from the end of the file
        If Not Right(FixStr(sHead), 1) = Chr(25) Then
            'this is not a compiled script file
            Execute1 s_Cmd 'Call Execute1
            Close #iFile
        Else
            iSize = CLng(Mid(sHead, 1, Len(sHead) - 1)) - 2 'Get the correct script filesize
            Buffer = Space(iSize) 'Create a buffer to hold the script
            Get #iFile, LOF(iFile) - (iSize + 14), Buffer 'Get the script data
            AddCode Buffer  'Add the code to the engine to be executed
            
            'Clear up
            ThisFile = ""
            Buffer = ""
            iSize = 0
            sHead = ""
            s_Cmd = ""
        Close #iFile
    End If
    
End Sub
