Attribute VB_Name = "ModResent"
Public Sub RemoveRecentFile(lzFile As String)
Dim List() As String, sLine As String, nFile As Long
Dim iCount As Integer, lnCnt As Integer, n As Integer

lnCnt = -1
nFile = FreeFile

    Open Recent_File For Input As #nFile
        Do While Not EOF(nFile)
            Input #nFile, sLine
            If Len(sLine) > 0 Then
                lnCnt = lnCnt + 1
                ReDim Preserve List(lnCnt)
                
                If LCase(lzFile) = LCase(sLine) Then
                    List(lnCnt) = ""
                Else
                    List(lnCnt) = sLine
                End If
                sLine = ""
            End If
            DoEvents
        Loop
    Close #nFile

    For n = 0 To lnCnt
        If Len(List(n)) > 0 Then
            sBuff = sBuff & List(n) & vbCrLf
        End If
    Next n

    Open Recent_File For Output As #nFile
        Print #nFile, sBuff
    Close #nFile
    sBuff = ""
    
End Sub
Public Sub WriteRecentList(lzFile As String)
Dim List() As String, sLine As String, nFile As Long
Dim iCount As Integer, lnCnt As Integer, n As Integer

sLine = ""
lnCnt = -1
nFile = FreeFile
Open Recent_File For Input As #nFile
    Do While Not EOF(nFile)
        Input #nFile, sLine
        If Len(sLine) > 0 Then
            lnCnt = lnCnt + 1
            ReDim Preserve List(lnCnt)
            List(lnCnt) = sLine
            sLine = ""
        End If
        DoEvents
    Loop
Close #nFile

If lnCnt >= RECENT_MAX_COUNT Then
    List(0) = lzFile
    GoTo WriteFile:
    Exit Sub
End If

If lnCnt = -1 Then
    lnCnt = 0
    ReDim Preserve List(lnCnt)
    List(lnCnt) = lzFile
    GoTo WriteFile:
Else
    lnCnt = UBound(List) + 1
    ReDim Preserve List(lnCnt)
    List(lnCnt) = lzFile
    GoTo WriteFile:
End If

WriteFile:

For n = 0 To lnCnt
    sBuff = sBuff & List(n) & vbCrLf
Next

Open Recent_File For Output As #nFile
    Print #nFile, sBuff
Close #nFile

sBuff = ""

End Sub

Public Function BuildRecentList(lzFile As String, mIncFile As String) As String
Dim nFile As Long, iCount As Integer
Dim StrA As String, sBuffer As String, StrB As String, StrC As String

    nFile = FreeFile
    sBuffer = OpenFile(mIncFile)
    
    LoadRecentList = ""
    frmMain.Toolbar1.Buttons(2).ButtonMenus.Clear
    
    Open lzFile For Input As #nFile
        Do While Not EOF(nFile)
            Input #nFile, StrA
            If Len(StrA) > 0 Then
                iCount = iCount + 1
                StrB = sBuffer
                StrB = Replace(StrB, "%File", GetFileName(StrA))
                StrB = Replace(StrB, "$File", StrA)
                StrC = StrC & StrB
                frmMain.Toolbar1.Buttons(2).ButtonMenus.Add , "b:" & iCount, StrA
            End If
            DoEvents
        Loop
    Close #nFile

    If Len(StrB) = 0 Then
        Exit Function
    Else
        BuildRecentList = StrC
        StrA = "": StrB = "": StrC = ""
    End If
    
End Function

Public Function IsInList(lzFile As String) As Boolean
Dim mList() As String, nFile As Long, sLine As String
    nFile = FreeFile
    
    IsInList = False
    
    Open Recent_File For Input As #nFile
        Do While Not EOF(nFile)
            Input #nFile, sLine
            If LCase(sLine) = LCase(lzFile) Then
                IsInList = True
                sLine = ""
                Exit Do
                Exit Function
            End If
            DoEvents
        Loop
    Close #nFile
    
End Function

Public Sub LoadHomePage(mWebB As WebBrowser)
Dim StrA As String, StrB As String
On Error Resume Next
    StrA = BuildRecentList(Recent_File, IncFile)
    StrB = OpenFile(StartPage)
    StrB = Replace(StrB, "<!--RecLst -->", StrA)
    SaveFile WebTemp, StrB
    mWebB.Navigate WebTemp
    StrA = "": StrB = ""
End Sub
