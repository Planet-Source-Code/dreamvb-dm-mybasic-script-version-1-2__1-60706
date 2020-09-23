Attribute VB_Name = "modOutput"
Sub PrintA(Optional LineFreed As Boolean = False)
Dim StrA As String
    'Print to the console
    
    StrA = Eval(ProcessLine) 'Get the String to be places on the console
    If isEmptyLine(StrA) Then
        'No expression was found so we abort
        Abort 8, CurrentLine, "PRINT", " = Expression"
        Exit Sub
    End If
    
    If Left(ProcessLine, 1) = "#" Then 'check for hash sign
        'Ok it looks like it may be a print statement for fileIO
        ProcessLine = RevStrLeft(ProcessLine, 1)
        GetIOMethod
    Else
        cWriteLine StrA
    End If
        
    StrA = ""
    
End Sub

Sub Locate()
Dim e_pos As Integer, A As Integer, B As Integer
On Error GoTo LocateErr:

    e_pos = CharPos(ProcessLine, ",") 'Get parm position
    
    If isEmptyLine(ProcessLine) Or e_pos = 0 Then
        'No expression was found so we abort
        Abort 8, CurrentLine, "LOCATE", "Expression,Expression"
        Exit Sub
    End If
    
    'Get both parms for the function
    A = CInt(Eval(Mid(ProcessLine, 1, e_pos - 1)))
    B = CInt(Eval(Mid(ProcessLine, e_pos + 1, Len(ProcessLine))))
    cSetCursorPosition A, B 'Position the text on the console
    A = 0: B = 0: e_pos = 0 ' Clean up
    
    Exit Sub
LocateErr:
    A = 0: B = 0: e_pos = 0
    If Err Then Abort 2, CurrentLine, Err.Description & " LOCATE " & ProcessLine
End Sub

Sub doColor()
Dim e_pos As Integer, A As Long, B As Long
    On Error GoTo ColorErr:
    e_pos = CharPos(ProcessLine, ",") 'Get parm position
    
    'Check that we don't have an empty line and the parms exsits
    If isEmptyLine(ProcessLine) Or e_pos = 0 Then
        Abort 8, CurrentLine, "COLOR", "Expression,Expression"
        Exit Sub
    End If
    'Extract the foreground and background color values
    
    A = CLng(Eval(Mid(ProcessLine, 1, e_pos - 1)))
    B = CLng(Eval(Mid(ProcessLine, e_pos + 1, Len(ProcessLine))))
    
    TextAttribute = A
    TextAttribute = B
    A = 0: B = 0: e_pos = 0 ' Clean up
    
    Exit Sub
ColorErr:
    If Err Then
        A = 0: B = 0: e_pos = 0
        If Err Then Abort 2, CurrentLine, Err.Description & " COLOR " & ProcessLine
    End If
End Sub
