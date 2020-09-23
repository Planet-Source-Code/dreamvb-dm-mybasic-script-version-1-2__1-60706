Attribute VB_Name = "modEval"
Function Token(Expression, pos) As Variant
Dim s_FuncName As String, e_pos As Long, var_Idx As Integer
Dim Value As Variant
    
    inSideQuotes = False
    Dim ch As String
    Dim pl As Integer, es As Integer
    
    Do
        ch = Mid(Expression, pos, 1)
            
        If isOperator(ch) Then
            Exit Do
        ElseIf ch = "(" Then
            pos = pos + 1
            pl = 1
            e_pos = pos
            'Look for the brackets in the expression
            Do
                ch = Mid(Expression, pos, 1)
                If ch = "(" Then pl = pl + 1
                If ch = ")" Then pl = pl - 1
                pos = pos + 1
            Loop Until pl = 0 Or pos > Len(Expression)
                
            Value = Mid(Expression, e_pos, pos - e_pos - 1) 'Get value of the function name
            s_FuncName = LCase(Trim(Token)) 'Get the function name
                
            If s_FuncName = "" Then
                'if no function name was found we just return the token
                Token = Eval(Value)
            End If
            
            'Built in functions
            Select Case s_FuncName
                'Strings
                Case "chr": Token = Chr(Eval(Value))
                Case "asc": Token = Asc(Eval(Value))
                Case "str": Token = Str(Eval(Value))
                Case "len": Token = Len(Eval(Value))
                Case "lcase": Token = LCase(Eval(Value))
                Case "ucase": Token = UCase(Eval(Value))
                Case "strreverse": Token = StrReverse(Eval(Value))
                Case "strconv": Token = DoStrFunction1(Value, 3)
                Case "strcomp": Token = DoStrComp(CStr(Value))
                Case "space": Token = Space(Eval(Value))
                Case "split": Token = DoStrFunction1(Value, 5)
                Case "string": Token = DoStrFunction1(Value, 2)
                Case "revnull": Token = RevNull(Eval(Value))
                Case "monthname": Token = DoStrFunction1(Value, 4, True)
                Case "mid": Token = DoMid(CStr(Value))
                Case "left": Token = DoStrFunction1(Value, 0)
                Case "right": Token = DoStrFunction1(Value, 1)
                'Maths
                Case "abs": Token = Abs(Eval(Value))
                Case "atn": Token = Atn(Eval(Value))
                Case "sin": Token = Sin(Eval(Value))
                Case "cos": Token = Cos(Eval(Value))
                Case "exp": Token = Exp(Eval(Value))
                Case "log": Token = Log(Eval(Value))
                Case "log10": Token = Log(Eval(Value)) \ Log(10)
                Case "rnd": Token = Rnd(Eval(Value))
                Case "int": Token = Int(Eval(Value))
                Case "isfloat": Token = isFloat(Eval(Value))
                Case "isnumeric": Token = IsNumeric(Eval(Value))
                Case "sqr": Token = Sqr(Eval(Value))
                Case "eval": Token = Eval(Value)
                Case "tan": Token = Tan(Eval(Value))
                'date/time
                Case "isdate": Token = Abs(IsDate((Eval(Value))))
                'FileIO
                Case "lof": Token = LOF(Eval(Value))
                Case "loc": Token = Loc(Eval(Value))
                Case "eof": Token = EOF(Eval(Value))
                Case "curdir": Token = CurDir(Eval(Value))
                Case "filedatetime": Token = FileDateTime(Eval(Value))
                Case "filelen": Token = FileLen(Eval(Value))
                Case "dir": Token = FileIOOperation(m_Dir, Value)
                Case "shell": Token = FileIOOperation(mShell, Value)
                Case "input": Token = FileIOOperation(mInput, Value)
                'Other
                Case "environ": Token = Environ(Eval(Value))
                Case "vartype": Token = GetVarType(CStr(Value))
                Case "lbound": Token = VarBound(Value, 0)
                Case "ubound": Token = VarBound(Value, 1)
                Case "array": Token = DoStrFunction1(Value, 6, True)
                Case Else
                    'Here we check for a variable as it may be an array
                    'At the moment it only support for varient as for the split function
                    var_Idx = VariableIndex(s_FuncName) 'Get the variable index
                    If Not (var_Idx = -1) And (GetVarType(s_FuncName) = nvar) Then
                        'if the variable index is found and of type varient then
                        ' Eval the data and return it
                        Token = Eval(mVarStack(var_Idx).VarData(Eval(Value)))
                    End If
            End Select
            
            ElseIf ch = Chr(34) Then
                    inSideQuotes = True 'Yes we are inside a string of quotes
                    pl = 1
                    pos = pos + 1
                    Do
                        ch = Mid(Expression, pos, 1)
                        pos = pos + 1
                        If ch = Chr(34) Then
                            If Mid(Expression, pos, 1) = Chr(34) Then
                                Value = Value & Chr(34)
                                pos = pos + 1
                            Else
                                Exit Do
                            End If
                        Else
                            Value = Value & ch
                        End If
                    Loop Until pl = 0 Or pos > Len(Expression)
                    Token = Value
            Else
                Token = Token & ch 'Return the token
                pos = pos + 1
            End If
            
   Loop Until pos > Len(Expression)
   
   Token = ReturnData(CStr(Token)) 'Return token
   
End Function

Function isOperator(StrExp As String) As Boolean
    isOperator = False
    If StrExp = "+" Or StrExp = "-" Or StrExp = "*" Or StrExp = "\" Or StrExp = "/" _
    Or StrExp = "&" Or StrExp = "^" Or StrExp = "=" Or StrExp = "<" Or StrExp = ">" Or StrExp = "%" Then _
    isOperator = True
End Function

Public Function Eval(Expression As Variant)
Dim iCounter As Integer, sOperator As String, Value As Variant
Dim iTmp As Variant, ch As String, ch2 As String

On Error Resume Next

    iCounter = 1
    
    Do While iCounter <= Len(Expression)
        ch = Mid(Expression, iCounter, 1)
        
        If isOperator(ch) Then
            sOperator = ch
            iCounter = iCounter + 1
        End If
        
        Select Case sOperator
            Case ""
                Value = Token(Expression, iCounter)
                'Only trim the line if it's a string
                If Not IsNumeric(Value) Then Value = Trim(Value)
            Case "+"
                Value = Value + Token(Expression, iCounter)
            Case "-"
                Value = Value - Token(Expression, iCounter)
            Case "*"
                Value = Value * Token(Expression, iCounter)
            Case "\"
                Value = Value \ Token(Expression, iCounter)
            Case "/"
                Value = Value / Token(Expression, iCounter)
            Case "%"
                Value = Value Mod Token(Expression, iCounter)
            Case "^"
                Value = Value ^ Token(Expression, iCounter)
            Case "<"
                Value = Abs(Value < Token(Expression, iCounter))
            Case ">"
                Value = Abs(Value > Token(Expression, iCounter))
            Case "&"
                Value = Value & Token(Expression, iCounter)
                If Right(Value, 1) = Chr(32) Then Value = Left(Value, Len(Value) - 1)
                If Left(Value, 1) = Chr(32) Then Value = Right(Value, Len(Value) - 1)
            Case "="
                iTmp = Token(Expression, iCounter)
                Value = Value = iTmp: Value = Abs(Value)
            Case Else
                Value = ""
        End Select
    Loop
    
    Eval = Value
    
    'Clear up
    Value = ""
    ch = ""
    
    If Err Then Abort 2, CurrentLine, Err.Description
    
End Function
