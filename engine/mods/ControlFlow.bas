Attribute VB_Name = "ControlFlow"
Sub DoForNext()
    'This function is used to process the loop
    If Eval(LoopStack.mExpr) = 0 Then
        'While expression is not true process the loop
        CurrentLine = LoopStack.mCurrentLine 'Keep within the loop
        LoopStack.mStart = GetVar(LoopStack.mVarName) + 1 'Add one to the loop counter
        SetVariableData LoopStack.mVarName, Int(LoopStack.mStart) 'Add the loop counter to the varstack
    End If

End Sub

Sub DoFor()
Dim e_pos As Integer
Dim lpVarName As String, var_Idx As Integer, sTmp As Variant, sTmp2 As Variant
Dim sHold1 As Variant, sHold2() As String, I As Integer, iSize As Integer

    iSize = -1

    If isEmptyLine(ProcessLine) Then Abort 9, CurrentLine
    ReDim sHold2(0) 'Resize
    sHold1 = Split(ProcessLine, " ") 'Split the line by it's spaces
    
    For I = 0 To UBound(sHold1)
        If Len(Trim(sHold1(I))) <> 0 Then
            iSize = iSize + 1
            ReDim Preserve sHold2(iSize)
            sHold2(iSize) = LCase(sHold1(I))
        End If
    Next
    
    Erase sHold1
    I = 0
    
    'Get the infomation need for out loop from the created array sHold2
    '
    var_Idx = VariableIndex(sHold2(0)) 'get the variable counter
    If var_Idx = -1 Then Abort 2, CurrentLine, "Required: variable"
    If Not sHold2(1) = "=" Then Abort 2, CurrentLine, "Required: '='" 'Get for assignment
    sTmp = Eval(sHold2(2)) 'Get start position of the loop
    If (Len(sTmp) = 0 Or Not IsNumeric(sTmp)) Then Abort 9, CurrentLine
    If Not sHold2(3) = "to" Then Abort 2, CurrentLine, "Required: TO" 'Locate TO
    
    'Get loop end postion
    sTmp2 = Eval(sHold2(4)) 'Get start position of the loop
    If (Len(sTmp2) = 0 Or Not IsNumeric(sTmp2)) Then Abort 9, CurrentLine
    
    LoopStack.mCurrentLine = CurrentLine 'Current executeing line
    LoopStack.mStart = CLng(sTmp) 'Stote the start postion
    LoopStack.mEnd = CLng(sTmp2) 'Stote the end postion
    LoopStack.mExpr = sHold2(0) & "=" & LoopStack.mEnd 'expression
    LoopStack.mVarName = sHold2(0) 'Variables name
    SetVariableData sHold2(0), LoopStack.mStart
    Erase sHold2
End Sub

Sub doGoto()
Dim sTmp_Label As String
Dim bFoundIdx As Integer
    'this function is used for the goto Statement
    ' at the moment it does seem to work quite well and does move between the lines.
    ' tho it may need some tweaks in up comming versions.
    
    'Below line check if the executeing ProcessLine is empty
    If isEmptyLine(ProcessLine) Then Abort 8, CurrentLine, "GOTO", "<Label>"
    
    'Below is used to locate the goto label in the script
    bFoundIdx = SerachList(0, ProcessLine)
    
    If bFoundIdx = -1 Then
        'If the label was not found we show the error
        Abort 2, CurrentLine, "Label not defined"
    Else
        'Move to the current line
        CurrentLine = bFoundIdx
    End If
    
End Sub

Sub DoSelect()
Dim Strln As String
Dim I As Long, EndSelectLn As Long, var_Idx As Integer
Dim e_pos As Integer, bFoundIdx As Boolean, lpVarName As String

Dim Case_Label As Variant ' Label used for Select case statement
Dim Case_Check_Label As Variant ' Label used for Select case statement
Dim Case_Line As Integer

    EndSelectLn = -1
    If isEmptyLine(ProcessLine) Then Abort 2, CurrentLine, "Required: Case"
    'Locate the end of select
    For I = CurrentLine To UBound(LineHolder)
        If LCase(Trim(LineHolder(I))) = "end select" Then
            EndSelectLn = I
        End If
    Next
    
    'Check that we found the end of select index
    If EndSelectLn = -1 Then Abort 2, CurrentLine, "Select Case without End Select"
    
    'Extract the variable name
    e_pos = CharPos(ProcessLine, " ")
    If e_pos = 0 Then Abort 2, CurrentLine, "Required: Case"
    
    lpVarName = Trim(Mid(ProcessLine, e_pos, Len(ProcessLine))) 'Extract variable name
    
    var_Idx = VariableIndex(lpVarName) 'Get varibale index
    If var_Idx = -1 Then Abort 6, CurrentLine, lpVarName 'No variable found on stack so we abort
    
    Case_Check_Label = Eval(GetVar(lpVarName)) 'Get the variables data
   
    For I = 0 To UBound(LineHolder) 'Loop tho all the script lines in LineHolder
        Strln = LCase(Trim(LineHolder(I))) 'Trim down the line removeing any white spaces
        If Left(Strln, 4) = "case" Then 'check for case
            e_pos = CharPos(Strln, " ") 'Get position of the space eg Case <space> expression
            Case_Label = Eval(Trim(Mid(Strln, e_pos + 1, Len(Strln) - 4))) 'Extract the expression
            If Len(Case_Label) = 0 Then Abort 9, CurrentLine 'check for an expression
            
            If Case_Label = Case_Check_Label Then
                'if Case_Label matchs Case_Check_Label then we set the line
                Case_Line = I 'Line number to move to
                bFoundIdx = True 'Yes we found a match
                Exit For
            ElseIf Case_Label = "else" Then 'Found a case else
                bFoundIdx = True
                Case_Line = I 'Line number to move to
                Exit For
            Else
                bFoundIdx = False 'No Match
            End If
        End If
    Next
    
    If bFoundIdx Then
        'Yes we have a match now move to that line
        CurrentLine = Case_Line
    Else
        'Move to the end of the end select
        CurrentLine = EndSelectLn
    End If
    
    'Clear up
    bFoundIdx = False
    Case_Label = ""
    Case_Check_Label = ""
    Case_Line = 0

End Sub

Sub DoCase()
    'This maynot be compleate yet
    CurrentLine = CurrentLine + 2
End Sub

Sub DoExit()
    Dim ln_Idx As Integer, sLine As String
    'Note this is still not yet finsihed
    ' If finsih this when I have added IF Support as it will serve a better purpose
    If isEmptyLine(ProcessLine) Then
        'it must be just a simple EXIT keyword
        CurrentLine = LineCount
        Exit Sub
    End If

    sLine = UCase(ProcessLine) 'Check for keywords eg EXIT FOR,EXIT DO
    
    Select Case sLine
        Case "FOR"
            ln_Idx = SerachList(CInt(CurrentLine), "next") 'Get the location of the next keyword
            If ln_Idx = -1 Then Abort 2, CurrentLine, "For without Next" ' next was not found so we abort
            CurrentLine = ln_Idx 'Move the current index of next
            sLine = "" 'Clean up
            Exit Sub
        Case Else
            Abort 1, CurrentLine, ProcessLine
    End Select
    
End Sub

Sub DoLoop()
Dim e_pos As Integer, sLoopType As String, sExp As String, I As Integer, sLine As String
    
    ' Well this is our code to deal with Do Loops, While Loops Loop Until
    ' Still needs to support the nesting of loops, I try and cover this in the next time
    ' but for now I think you find it usfull for basic things

    e_pos = CharPos(ProcessLine, " ") 'Find the first space
    
    If isEmptyLine(ProcessLine) And e_pos = 0 Then
        sExp = vbNullChar
        For I = CurrentLine To LineCount 'Loop tho all the script lines
            sLine = LCase(Trim(LineHolder(I))) 'Get current line
            'Now we check for Loop as the end of loop marker
            If sLine = "loop" Then
                sExp = sLine 'Get the expression
                sLine = "" 'Clear
                Exit For
            Else
                e_pos = CharPos(sLine, " ")
                If Trim(Left(sLine, e_pos)) = "loop" Then
                    sLine = ""
                    sExp = LineHolder(I)
                    Exit For
                Else
                    sExp = vbNullChar
                    sLine = ""
                End If
            End If
        Next
        
        If Len(sExp) = 0 Or sExp = vbNullChar Then Abort 2, CurrentLine, "Required: Loop"
        
        If sExp = "loop" Then
            '
            Exit Sub
        Else
            'it Looks like a Loop Until
            sLine = Trim(Right(sExp, Len(sExp) - 4))
            e_pos = CharPos(sLine, " ")
            If e_pos = 0 Then Abort 9, CurrentLine
            If Not LCase(Trim(Left(sLine, e_pos))) = "until" Then Abort 2, CurrentLine, "Required: Loop"
            'Fix expression by removeing until keyword
            sLine = Trim(Right(sLine, Len(sLine) - 5))
            If sLine = "" Then Abort 9, CurrentLine 'No expression found so exit
            'Stote the expression
            sExp = sLine
            LoopStack.mCurrentLine = CurrentLine
            LoopStack.mExpr = sExp
            
            e_pos = 0
            sExp = ""
            sLine = ""
            Exit Sub
        End If
    Else
        sLoopType = Trim(UCase(Left(ProcessLine, e_pos))) 'Extract Loop type
        If e_pos = 0 Then Abort 9, CurrentLine 'Make sure we have a expression
        sExp = Trim(Mid(ProcessLine, e_pos, Len(ProcessLine)))
    End If
    
    If sLoopType = "WHILE" Then
        'Do while loop
        'locate the end of the loop]
        If SerachList(CInt(CurrentLine), "loop") = -1 Then Abort 2, CurrentLine, "Required: Loop"
        LoopStack.mCurrentLine = CurrentLine 'Current line to
        LoopStack.mExpr = sExp 'Expression
        Exit Sub
    ElseIf sLoopType = "UNTIL" Then
        'Do until
        If SerachList(CInt(CurrentLine), "loop") = -1 Then Abort 2, CurrentLine, "Required: Loop"
        LoopStack.mCurrentLine = CurrentLine 'Current line to
        LoopStack.mExpr = sExp 'Expression
        Exit Sub
    Else
        e_pos = 0: sLoopType = "": sExp = ""
        Abort 2, CurrentLine, "Required: While or Until or end of statement"
    End If
    
End Sub

Sub DoEndOfLoop()
    'This is used to process the loop until the expresion is met
    If Eval(LoopStack.mExpr) = 0 Then
        CurrentLine = LoopStack.mCurrentLine
    End If
End Sub
