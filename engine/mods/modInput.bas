Attribute VB_Name = "modInput"
Sub DoClearVars()
Dim VarLst As Variant, I As Integer, idx As Integer
Dim sTmp As String

    ' Now this DoClearVars Function is used to reset a variables state
    ' it's the same as using StrVar = "" or NumVar = 0
    ' the only different is that it can be used to clear all variables with one pass
    ' eg using the CLEAR keyword in your code will reset all variables
    ' while the entanded feature CLEAR Var1,Var2 will only clear what's in the list
    sTmp = ""
    
    If isEmptyLine(ProcessLine) Then
        'No variable list was found we clear evey variable
        ResetAllVars
        Exit Sub
    Else
        VarLst = Split(ProcessLine, ",")
        
        For I = 0 To UBound(VarLst)
            sTmp = Trim(VarLst(I))
            idx = VariableIndex(sTmp) 'Get the variables index
            If idx = -1 Then 'Variable not found
                Erase VarLst 'Erase the variable list
                Abort 6, CurrentLine, sTmp 'abort
            Else
                ResetVariable idx 'call ResetVariable
                sTmp = ""
            End If
        Next
        Erase VarLst
    End If
    
End Sub

Sub DoConst()
Dim e_Pos As Integer
Dim sTmp As String, sConstName As String, ConstData As Variant

    If isEmptyLine(ProcessLine) Then Abort 2, CurrentLine, _
    "Required: identifier" & vbCrLf & "USE: CONST=<expression>"
    
    sTmp = Trim(ProcessLine)
    
    e_Pos = CharPos(sTmp, "=") 'Get the assign position
    If e_Pos = 0 Then Abort 2, CurrentLine, "Required: '='"
    
    sConstName = Trim(Mid(sTmp, 1, e_Pos - 1)) 'Extract the const name
    'Check that the const name is vaild
    If Not isVaildVarName(sConstName) Then Abort 2, CurrentLine, "Inviald variable name: " & sConstName
    
    'Next we need to check if the const name is the same as a variable name on the stack
    ' if it is we can't add it becuase this will be a Duplicateed name
    If VariableIndex(sConstName) <> -1 Then Abort 2, CurrentLine, "Duplicate declaration in current scope"
    'Now we need to check if the consts is within the const stack
    ' if it is then it can not be added. unless it a system const that can be chnaged.
   If ConstIndex(sConstName) <> -1 Then Abort 2, CurrentLine, "Assignment to constant not permitted"
   ConstData = Eval(Trim(Mid(sTmp, e_Pos + 1, Len(sTmp))))
   
   If IsEmpty(ConstData) Then Abort 9, CurrentLine 'Check that the const expression is not empty
   AddConst Trim(sConstName), ConstData, False 'We can now add our new const
   
   'Clear up
   e_Pos = 0
   sConstName = ""
   ConstData = ""
   sTmp = ""
   
End Sub

Sub DoDim(lpstr As String)
Dim StrA As String, e_Pos As Integer, n_pos As Integer, StrVarName As String, nVarType As VarType
    
    'Updates made:
    ' Checking of vaild variable names
    ' checking for const variable names
    
    If isEmptyLine(lpstr) Then Abort 4, CurrentLine
    e_Pos = CharPos(lpstr, Chr(32))
    
    If e_Pos = 0 Then
        StrVarName = Trim(lpstr) 'Get variable name
        'Line below check if the variable is a const
        If ConstIndex(StrVarName) <> -1 Then Abort 2, CurrentLine, "Duplicate declaration in current scope"
        
        'Check for a vaild variable name
        If Not isVaildVarName(StrVarName) Then
            Abort 2, CurrentLine, "Inviald variable name: " & StrVarName
        End If
        
        'check if the variable is not already in the stack
        If VariableIndex(StrVarName) <> -1 Then
            'Variable was already found
            e_Pos = 0
            Abort 5, CurrentLine, StrVarName
        Else
            'Add the variable to the variables stack
            AddVariable StrVarName, nvar, , , ""
            e_Pos = 0
            Exit Sub
        End If
    Else
        'Get the variable name
        StrVarName = Trim(Mid(lpstr, 1, e_Pos - 1))
        
        'Check for a vaild variable name
        If isVaildVarName(StrVarName) = False Then
            Abort 2, CurrentLine, "Inviald variable name: " & StrVarName
        End If
        
        'Check for the variables datatype
        n_pos = InStr(e_Pos + 1, lpstr, Chr(32), vbBinaryCompare)
        If n_pos = 0 Then
            Abort 2, CurrentLine, "Required: DataType"
            e_Pos = 0: n_pos = 0: StrVarName = ""
            Exit Sub
        End If
        
        'Make sure that we have AS in the expression
        StrA = UCase(Trim(Mid(lpstr, e_Pos + 1, n_pos - e_Pos - 1)))
        If StrA <> "AS" Then
            StrA = "": e_Pos = 0: StrVarName = ""
            Abort 2, CurrentLine, "Required: AS"
        End If
        
        StrA = Trim(Mid(lpstr, n_pos + 1, Len(lpstr))) 'Extract the variables datatype
        nVarType = GetVarTypeFromStr(StrA) 'Store the varibales datatype
        
        If nVarType = NoKnownErr Then Abort 7, CurrentLine, StrA 'invaild datatype
        StrA = ""
        
        ' check that the variable is in the stack
         If VariableIndex(StrVarName) <> -1 Then
            e_Pos = 0: n_pos = 0
            Abort 5, CurrentLine, StrVarName
        Else
            'We have our variable and it's data type so we can add it to the variable stack.
            AddVariable StrVarName, nVarType, False, , SetVarDefault(nVarType)
            e_Pos = 0: n_pos = 0: StrA = "": StrVarName = ""
        End If
    End If
    
End Sub

Sub DoAssign1(lpExpr As String, LetAssign As Boolean, Optional AssignVarName As String)
Dim e_Pos As Integer, StrVarName As String, AssignData As Variant
Dim iTemp As Variant
    'Ok now our assign sub can now deal with two assignments:
    ' the LET assign that we original had and also normal assignments such as A = B + C
    ' not bad at all and we not had to use a new sub and only about 4 lines of changes
    
    If isEmptyLine(lpExpr) Then
        If LetAssign Then 'This tests if we are dealing with a LET assign
            Abort 8, CurrentLine, "LET", " = Expression"
        End If
    End If
    
    'Check for the assign pos
    e_Pos = CharPos(lpExpr, "=") 'Get location of the assignment sign
    If e_Pos = 0 Then Abort 2, CurrentLine, "Required: '='" 'check for the assignment sign
    
    'Edited code
    If LetAssign Then
        'we use the for the LET assign keyword
        StrVarName = Trim(Mid(lpExpr, 1, e_Pos - 1)) 'Extract the variable name
    Else
        'For a normal assign we do it slity different
        StrVarName = AssignVarName 'Assign the variable name from AssignVarName to StrVarName
    End If
    
    CurrentVar = StrVarName 'new
    
    'check that the variable name above is in the variable stack
    If VariableIndex(StrVarName) = -1 Then
        StrVarName = ""
        Abort 6, CurrentLine, StrVarName
    Else
        AssignData = Trim(Mid(lpExpr, e_Pos + 1, Len(lpExpr))) 'Extract the expression
        
        If AssignData = "" Then
            'No expression was found
            StrVarName = "": AssignData = ""
            Abort 9, CurrentLine
        Else
            iTemp = Eval(AssignData) 'eval the assign data
            If iTemp <> vbNullChar Then 'new code
                SetVariableData StrVarName, SetVarDataType(GetVarType(StrVarName), iTemp)
            End If
        End If
    End If
End Sub

Sub DoInput()
Dim lpVarName As String, Str_Tmp As String
    e_Pos = CharPos(ProcessLine, ",") 'Find the position of the parm seprator ,
    
    If isEmptyLine(ProcessLine) Or e_Pos = 0 Then
        'No expression was found so we abort
        Abort 8, CurrentLine, "INPUT", "[Prompt],Variable"
        Exit Sub
    End If
    
    'Extract the variable name
    lpVarName = Trim(Mid(ProcessLine, e_Pos + 1, Len(ProcessLine)))
    'Check that the variable is in the variables stack
    If VariableIndex(lpVarName) = -1 Then
        'Variable was not found so we abort
        Abort 4, CurrentLine, lpVarName
        lpVarName = ""
        Exit Sub
    Else
        Str_Tmp = Eval(Mid(ProcessLine, 1, e_Pos - 1)) 'Extract the propmt message
        'Now we need to print the propmting message to the user
        
        If Left(ProcessLine, 1) = "#" Then
            Str_Tmp = ""
            ProcessLine = RevStrLeft(ProcessLine, 1)
            DoFileInput
        Else
            cWriteLine Str_Tmp
            Str_Tmp = ""
            'Now we will use the console read command to get input form the user
            Str_Tmp = cReadConsole()
            'Stote the user input data into the variable ->lpVarName
            SetVariableData lpVarName, SetVarDataType(GetVarType(lpVarName), Str_Tmp)
            'Clean up used varaibles
            Str_Tmp = ""
            lpVarName = ""
            e_Pos = 0
        End If
    End If
End Sub

