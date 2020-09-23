Attribute VB_Name = "ModStrings"
'Mod file for dealing with our string functions

Function isFloat(lpVar As Variant) As Integer
    'Returns 1 is a number is float eg 1.5 .3
    If Not IsNumeric(lpVar) Then isFloat = 0: Exit Function
    isFloat = CharPos(CStr(lpVar), ".") <> 0
End Function

Function DoStrComp(lpData As String) As String
Dim e_pos As Integer, n_pos As Integer
Dim TempHold(2) As String
On Error GoTo StrErr:
    
    'Used to compare two strings
    
    e_pos = StrFind(0, lpData, ", ")
    n_pos = StrFind(e_pos, lpData, ", ")
    If (e_pos = 0) Then Abort 11, CurrentLine, "StrComp", "StrComp String1, String2, [Compare]"
    TempHold(0) = Eval(Trim(Left(lpData, e_pos - 1))) 'Get the string part
    
    If n_pos = 0 Then
        TempHold(1) = Eval(Mid(lpData, e_pos + 1, Len(lpData)))
        TempHold(2) = vbNullChar
        DoStrComp = StrComp(TempHold(0), TempHold(1))
    Else
        TempHold(1) = Eval(Trim(Mid(lpData, e_pos + 1, n_pos - e_pos - 1)))
        TempHold(2) = Eval(Trim(Mid(lpData, n_pos + 1, Len(lpData))))
        DoStrComp = StrComp(TempHold(0), TempHold(1), TempHold(2))
    End If
    
Clean:
    e_pos = 0: n_pos = 0
    Erase TempHold
    
    Exit Function
StrErr:
    If Err Then Abort 2, CurrentLine, Err.Description
    
End Function

Function DoMid(lpData As String) As String
Dim e_pos As Integer, n_pos As Integer
Dim TempHold(2) As String
On Error GoTo StrErr:

    e_pos = StrFind(0, lpData, ", ")
    n_pos = StrFind(e_pos, lpData, ", ")
    If (e_pos = 0) Then Abort 11, CurrentLine, "Mid", "Mid String, Start, [Length]"
    TempHold(0) = Eval(Trim(Left(lpData, e_pos - 1))) 'Get the string part
    
    If n_pos = 0 Then
        TempHold(1) = Eval(Mid(lpData, e_pos + 1, Len(lpData)))
        TempHold(2) = vbNullChar
        DoMid = Mid(TempHold(0), TempHold(1)) 'Do mid only String and Start
    Else
        TempHold(1) = Eval(Trim(Mid(lpData, e_pos + 1, n_pos - e_pos - 1)))
        TempHold(2) = Eval(Trim(Mid(lpData, n_pos + 1, Len(lpData))))
        DoMid = Mid(TempHold(0), TempHold(1), TempHold(2)) 'Do mid String,Start,length
    End If
    
Clean:
    e_pos = 0: n_pos = 0
    Erase TempHold
    
    Exit Function
StrErr:
    If Err Then Abort 2, CurrentLine, Err.Description
    
End Function

Function DoStrFunction1(lValue As Variant, mOption As Integer, Optional isOptionl As Boolean = False)
Dim TempHold(2) As String
On Error GoTo StrErr:

    GetTwoParms CStr(lValue), TempHold, isOptionl 'Get the parms
    
    If mOption = 0 Then 'left
        DoStrFunction1 = Left(TempHold(0), TempHold(1))
    ElseIf mOption = 1 Then 'Right
        DoStrFunction1 = Right(TempHold(0), TempHold(1))
    ElseIf mOption = 2 Then 'string
        DoStrFunction1 = String(TempHold(0), TempHold(1))
    ElseIf mOption = 3 Then 'StrConv
        DoStrFunction1 = StrConv(TempHold(0), TempHold(1))
    ElseIf mOption = 4 Then 'MonthName
        If TempHold(1) = vbNullChar Then DoStrFunction1 = MonthName(TempHold(0)) Else DoStrFunction1 = MonthName(TempHold(0), TempHold(1))
    ElseIf mOption = 5 Then 'Split
        SetVariableData CurrentVar, Split(CStr(TempHold(0)), TempHold(1))
        DoStrFunction1 = vbNullChar 'Do not assign anything back to the variable
    ElseIf mOption = 6 Then 'Array
        SetVariableData CurrentVar, Split(lValue, ",")
        DoStrFunction1 = vbNullChar 'Do not assign anything back to the variable
    End If

CleanUp:
Erase TempHold

Exit Function

StrErr:
    If Err Then Abort 2, CurrentLine, Err.Description
    
End Function
