Attribute VB_Name = "Variables"
'OK this is the main mod were functions and Subs are kept/
' that are used for dealing with the variables and consts

'Variable datatypes
Enum VarType
    NoKnownErr = 0
    nString = 1
    nInteger = 2
    nvar = 3
    nlong = 4
    nDouble = 5
    nBoolean = 6
    nByte = 7
    nCurrency = 8
End Enum

'Variable stack
Private Type VarStack
    VariableName As String ' The Variables name
    VariableType As VarType ' Type of variable
    VarData As Variant ' Variables data
    isGlobal As Boolean 'Not used in this version
    CanChange As Boolean
End Type

'Consts stack
Private Type VarConsts
    ConstName As String
    isSystemConst As Boolean
    ConstsData As Variant
End Type

Public MaxVars As Long ' Hold the current number of variables
Public ConstMax As Long 'Hold the the numner of consts

Public mVarStack() As VarStack
Public mConstStack() As VarConsts

Public Function ConstIndex(lpName As String) As Integer
Dim idx As Integer, x As Integer
    
    'Locate an consts index based on it's name
    ' If lpName does not match mConstStack(x).ConstName then
    ' the default is passed
    idx = -1
    If ConstMax = -1 Then ConstIndex = idx: Exit Function
    
    For x = 0 To UBound(mConstStack) 'Lopp tho the variable stack
        If LCase(lpName) = mConstStack(x).ConstName Then
            'We have a match so store the consts index
            idx = x '--< store this
            Exit For 'Get out of this loop
        End If
    Next
    
    ConstIndex = idx 'Return idx
    
End Function

Public Function VariableIndex(lpVarName As String) As Integer
Dim idx As Integer, x As Integer
    
    'Locate an variables index based on it's name
    ' If lpVarName does nopt match mVarStack(x).VariableName then
    ' the default error index is returned -1
    idx = -1
    If MaxVars = -1 Then VariableIndex = idx: Exit Function
    
    For x = 0 To UBound(mVarStack) 'Lopp tho the variable stack
        If LCase(lpVarName) = mVarStack(x).VariableName Then
            'We have a match so store the variables index
            idx = x '--< store this
            Exit For 'Get out of this loop
        End If
    Next
    
    VariableIndex = idx 'Return good result index
    
End Function

Public Sub AddVariable(lVarName As String, lVarType As VarType, Optional isGlobalEx As Boolean = True, _
Optional isReadOnly As Boolean = False, Optional lpVarData As Variant)
    
    MaxVars = MaxVars + 1 'Keep a count of the total variables we have
    ReDim Preserve mVarStack(MaxVars) 'Resize the variable stack
    'Now we fill in the information for the current variable been added
    mVarStack(MaxVars).VariableName = LCase(lVarName)
    mVarStack(MaxVars).CanChange = isReadOnly ' Means if it can be changed by the user
    mVarStack(MaxVars).isGlobal = isGlobalEx ' Means can this variable be accessed outside of it's scope
    mVarStack(MaxVars).VarData = SetVarDataType(lVarType, lpVarData)  ' Get and set the variables data
    mVarStack(MaxVars).VariableType = lVarType ' Get the varibales datatype
End Sub

Public Sub AddConst(lpConstName As String, lpConstDat As Variant, lpIsSystem As Boolean)
    ConstMax = ConstMax + 1 'Keep a count of const we have
    ReDim Preserve mConstStack(ConstMax) 'Resize the consts stack
    'Now we fill in the information for the current const been added
    mConstStack(ConstMax).ConstName = LCase(lpConstName)
    mConstStack(ConstMax).isSystemConst = lpIsSystem
    mConstStack(ConstMax).ConstsData = lpConstDat
End Sub

Public Sub SetConst(lpConstIdx As Integer, lpData As Variant)
    mConstStack(lpConstIdx).ConstsData = lpData
End Sub


Public Sub SetVariableData(lVarName As String, Optional VarData As Variant)
Dim idx As Integer
    ' this function is used to set the variables data
    idx = VariableIndex(lVarName)
    mVarStack(idx).VarData = VarData
End Sub

Public Function GetVar(lpName As String) As Variant
Dim idx As Integer
    ' Function that is used to return the data from a Given variable
    ' if idx is returned -1 then a nullstring is sent back.
    idx = VariableIndex(lpName)
    
    If idx <> -1 Then
        GetVar = mVarStack(idx).VarData
        ' Return the variables data
        Exit Function
    End If
    
End Function

Public Function GetConst(lpName As String) As Variant
Dim idx As Integer, v As Variant
    
    ' Function that is used to return the data from a Given const
    ' if idx is returned -1 then a nullstring is sent back.
    idx = ConstIndex(lpName)
    
    If idx <> -1 Then
        If LCase(lpName) = "rnd" Then
            GetConst = Rnd
        ElseIf LCase(lpName) = "time" Then
            GetConst = Time
        ElseIf LCase(lpName) = "date" Then
            GetConst = Date
        ElseIf LCase(lpName) = "freefile" Then
            GetConst = FreeFile
        ElseIf LCase(lpName) = "dir" Then
            GetConst = Dir
        Else
            GetConst = mConstStack(idx).ConstsData
        ' Return the const data
        End If
        
        Exit Function
    End If
    
End Function

Function GetVarType(lpName As String) As VarType
    Dim idx As Integer
    'Returns a variables data type
    idx = VariableIndex(lpName)
    If idx <> -1 Then
        GetVarType = mVarStack(idx).VariableType
        Exit Function
    Else
        GetVarType = NoKnownErr
    End If
End Function

Function SetVarDataType(lpVarType As VarType, lpVarData As Variant) As Variant
On Error GoTo SetDataErr:
    ' This function is used for seting the proper datatypes with there data
    ' it also a good way to test for error such as overflows or incorrect datatypes
    Select Case lpVarType
        Case nInteger: SetVarDataType = CInt(lpVarData): Exit Function
        Case nString: SetVarDataType = CStr(lpVarData): Exit Function
        Case nvar: SetVarDataType = CVar(lpVarData): Exit Function
        Case nlong: SetVarDataType = CLng(lpVarData): Exit Function
        Case nDouble: SetVarDataType = CDbl(lpVarData): Exit Function
        Case nByte: SetVarDataType = CByte(lpVarData): Exit Function
        Case nCurrency: SetVarDataType = CCur(lpVarData): Exit Function
        Case nBoolean: SetVarDataType = Int(CBool(lpVarData)): Exit Function
        Case Else: Abort 3, CurrentLine
    End Select
    
    Exit Function
SetDataErr:
    Abort 2, CurrentLine, Err.Description
    
End Function

Function SetVarDefault(nType As VarType) As Variant
    'Set the default data of a new variable
    Select Case nType
        Case nInteger, nlong, nDouble, nBoolean, nByte, nCurrency: SetVarDefault = 0
        Case nString, nvar: SetVarDefault = ""
    End Select
    
End Function

Function GetVarTypeFromStr(lpVarType As String) As VarType
    'This function is used to return the numric value of a variables datatype
    ' the function works by checking the string vartype and returning the value
    ' also see ENUM VarType
    
    Select Case UCase(lpVarType)
        Case "STRING": GetVarTypeFromStr = nString ' String variable
        Case "INTEGER": GetVarTypeFromStr = nInteger ' Numberic variable
        Case "VARIANT": GetVarTypeFromStr = nvar 'Variant datatype
        Case "LONG": GetVarTypeFromStr = nlong 'Long Numberic
        Case "DOUBLE": GetVarTypeFromStr = nDouble
        Case "BYTE": GetVarTypeFromStr = nByte
        Case "BOOLEAN": GetVarTypeFromStr = nBoolean ' TRUE/FALSE
        Case "CURRENCY": GetVarTypeFromStr = nCurrency 'Floating point numbers
        Case Else: GetVarTypeFromStr = NoKnownErr 'Type not supported
    End Select
    
End Function

Public Sub ResetAllVars()
Dim I As Integer
    'Resets all variables on the stack to there default values
    If MaxVars = -1 Then Exit Sub
    For I = 0 To UBound(mVarStack)
        mVarStack(I).VarData = SetVarDefault(mVarStack(I).VariableType)
    Next
    
End Sub

Public Sub ResetVariable(lpVarIdx As Integer)
    'Reset a variable to it's default value
    mVarStack(lpVarIdx).VarData = SetVarDefault(mVarStack(lpVarIdx).VariableType)
End Sub

Function VarBound(lpVarName As Variant, lpBoundOp As Integer) As Integer
Dim idx As Integer, lpName As String
On Error GoTo BoundErr:
    lpName = CStr(lpVarName)
    
    idx = VariableIndex(lpName) 'Get variables index
    If idx = -1 Then Abort 2, CurrentLine, "Required: Variable"
    
    'Find out the variables type
    If GetVarType(lpName) = nvar Then
        'Varient
        If lpBoundOp = 0 Then
            'Get lower bound
            VarBound = LBound(mVarStack(idx).VarData)
        Else
            'Get upper bound
            VarBound = UBound(mVarStack(idx).VarData)
        End If
    Else
        'Not finsihed this will be for arrays
        ' so we just send an error
        Abort 2, CurrentLine, "Type Mistake"
    End If
    
    
    Exit Function
BoundErr:
    If Err Then Abort 2, CurrentLine, Err.Description
End Function
