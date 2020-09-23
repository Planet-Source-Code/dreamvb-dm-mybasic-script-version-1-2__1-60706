VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM MyBasic-Script Live-Update"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "frmLiveUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin dmLiveUpdate.DmDownload DmDownload1 
      Left            =   5550
      Top             =   3855
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   1530
      TabIndex        =   18
      Top             =   5250
      Width           =   1260
   End
   Begin VB.CommandButton cmdfolder 
      Caption         =   "...."
      Height          =   390
      Left            =   5565
      TabIndex        =   17
      Top             =   4680
      Width           =   435
   End
   Begin VB.TextBox txtoutput 
      Height          =   390
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "C:\"
      Top             =   4680
      Width           =   5400
   End
   Begin VB.ListBox lstUpdates 
      Height          =   1425
      Left            =   105
      TabIndex        =   7
      Top             =   2160
      Width           =   7005
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Get Update"
      Enabled         =   0   'False
      Height          =   345
      Left            =   105
      TabIndex        =   5
      Top             =   5250
      Width           =   1260
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Now"
      Height          =   315
      Left            =   1755
      TabIndex        =   4
      Top             =   1185
      Width           =   1140
   End
   Begin VB.ComboBox cboMirror 
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   1200
      Width           =   1485
   End
   Begin VB.PictureBox PicTop 
      Align           =   1  'Align Top
      BackColor       =   &H00E7E3DE&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   7140
      TabIndex        =   0
      Top             =   0
      Width           =   7140
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM MyBasic-Script Live-Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   750
         TabIndex        =   1
         Top             =   180
         Width           =   3810
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmLiveUpdate.frx":08CA
         Top             =   135
         Width           =   480
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Save update to:"
      Height          =   195
      Left            =   105
      TabIndex        =   15
      Top             =   4425
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   765
      Y1              =   4260
      Y2              =   4260
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   765
      Y1              =   4275
      Y2              =   4275
   End
   Begin VB.Label lblUrl 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2940
      TabIndex        =   14
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1725
      TabIndex        =   13
      Top             =   3975
      Width           =   45
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   165
      TabIndex        =   12
      Top             =   4005
      Width           =   45
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2940
      TabIndex        =   11
      Top             =   3705
      Width           =   405
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1725
      TabIndex        =   10
      Top             =   3705
      Width           =   435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   9
      Top             =   3705
      Width           =   720
   End
   Begin VB.Label lblupdates 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Updates"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label lblwait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2985
      TabIndex        =   6
      Top             =   1245
      Width           =   45
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   765
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   765
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mirrors:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   915
      Width           =   765
   End
   Begin VB.Line ln 
      BorderColor     =   &H00808080&
      X1              =   15
      X2              =   780
      Y1              =   780
      Y2              =   780
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Dim isDownload As Boolean
Dim Mirros() As String
Dim sTmp As Integer
Dim UpdateList As String
Dim DataList() As String

Private Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        Offset = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, Offset - 1)
    End If
End Function

Sub ShowUpdateList()
Dim nFile As Long, sLine As String, e_pos As Integer
Dim iSize As Integer

    'Open and show the updates information
    nFile = FreeFile
    Erase DataList
    lstUpdates.Clear
    
    Open UpdateList For Input As #nFile
        Do While Not EOF(nFile)
            Input #nFile, sLine
            e_pos = InStr(1, sLine, "update=", vbTextCompare)
            
            If e_pos <> 0 Then
                sLine = Right(sLine, Len(sLine) - 7)
                iSize = iSize + 1
                ReDim Preserve DataList(iSize)
                e_pos = InStr(1, sLine, "|")
                If e_pos <> 0 Then
                    lstUpdates.AddItem Left(sLine, e_pos - 1)
                End If
                
                DataList(iSize) = sLine
                sLine = ""
            End If
        Loop
    Close #nFile
    Kill UpdateList
    iSize = 0
    e_pos = 0
    
End Sub

Function GetFileName(lzUrlFile As String, Optional BySlash As String = "/") As String
Dim x As Integer, e_pos As Integer
    For x = 1 To Len(lzUrlFile)
        If Mid$(lzUrlFile, x, 1) = BySlash Then e_pos = x
    Next x
    
    If e_pos = 0 Then GetFileName = lzUrlFile: Exit Function
    
    GetFileName = Mid$(lzUrlFile, e_pos + 1, Len(lzUrlFile))
    
End Function

Function FixPath(lzPath As String) As String
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Private Sub OpenMirrros(lzFile As String)
Dim TFile As Long, sLine As String, e_pos As Integer, n_pos As Integer
    Dim MirrorName As String, x As Integer
    ' This sub is used to Parse all the information from the mirros file
    
    cboMirror.Clear 'Clear combo box
    TFile = FreeFile 'Get a free file Number
    
    If Not IsFileHere(lzFile) Then
        MsgBox "Unable to load Mirros List.", vbCritical, "File Not Found"
        Exit Sub
    Else
        Open lzFile For Input As #TFile
        Do While Not EOF(TFile)
            'Loop Though the file and locate the data
            Input #1, sLine 'Input each line
            e_pos = InStr(1, sLine, "=", vbBinaryCompare) 'Check = sign
            n_pos = InStr(e_pos + 1, sLine, "|", vbBinaryCompare) ' check for |
            If (e_pos > 0) And (n_pos > 0) Then 'Have we got a match of above
                If LTrim(UCase(Left(sLine, e_pos - 1))) = "MIRROR" Then
                    x = x + 1 'Inc counter
                    ReDim Preserve Mirros(x) 'Resize array
                    MirrorName = Trim(Mid(sLine, e_pos + 1, n_pos - e_pos - 1)) 'Extract mirror name
                    Mirros(x) = Trim(Mid(sLine, n_pos + 1, Len(sLine))) ' Extract mirror URL
                    cboMirror.AddItem MirrorName 'add mirror name to combo box
                End If
            End If
            DoEvents 'let us things process
        Loop
    End If
    
    cboMirror.Enabled = cboMirror.ListCount > 0
    cmdCheck.Enabled = cboMirror.Enabled
    ' Enable combo box and Check Now button, if we have a none zero count
    If cboMirror.ListCount > 0 Then cboMirror.ListIndex = 0
    
End Sub

Private Sub cboMirror_Change()
    cboMirror.ListIndex = sTmp
    'This stops the deleteion of of items from the combobox and stores the index
End Sub

Private Sub cboMirror_Click()
    'This stops the deleteion of of items from the combobox and sets the index to it's last one
    sTmp = cboMirror.ListIndex
End Sub

Private Sub cmdCheck_Click()
Dim Idx As Integer, sBuffer As String
    If UBound(Mirros) = 0 Then Exit Sub
    Idx = (sTmp + 1) 'Get the combo list index
    lblwait.Caption = "Please wait receiving update information"
    DmDownload1.DownloadFile Mirros(Idx), UpdateList
End Sub

Private Sub cmdDown_Click()
Dim lFileName As String
    isDownload = True
    cmdDown.Enabled = False
    lFileName = txtoutput.Text & GetFileName(lblUrl.Caption, "/")
    DmDownload1.DownloadFile lblUrl.Caption, lFileName, vbAsyncTypeByteArray
End Sub

Private Sub cmdExit_Click()
    If isDownload Then
        If MsgBox("There is still a download in process." _
        & vbCrLf & "Are you sure you want to exit now?", vbYesNo Or vbQuestion) = vbNo Then
            Exit Sub
        Else
            Unload frmMain
        End If
    Else
        Unload frmMain
    End If
    
End Sub

Private Sub cmdfolder_Click()
Dim sFolder As String
    sFolder = GetFolder(Me.hWnd, "Select Folder:")
    If Len(sFolder) = 0 Then Exit Sub
    
    txtoutput.Text = FixPath(sFolder)
    
End Sub

Private Sub DmDownload1_DownloadComplete(mCurBytes As Long, mMaxBytes As Long, LocalFile As String)
    If isDownload Then
        lblwait.Caption = "Download is now complete."
        isDownload = False
        cmdDown.Enabled = True
    Else
        If IsFileHere(UpdateList) Then
            ShowUpdateList
            cmdCheck.Enabled = True
            lblwait.Caption = "Update list received."
        End If
    End If
End Sub

Private Sub DmDownload1_DownloadProgress(mCurBytes As Long, mMaxBytes As Long)
    If isDownload Then
        'Display a wait message
        lblwait.Caption = "Please wait downloading update."
    End If
End Sub

Private Sub DmDownload1_LastError(StatusCode As Long, Status As String, LocalFile As String)
    MsgBox "There was an error while downloading the update." & vbCrLf _
    & vbCrLf & "Error Code: " & vbCrLf & Status, vbCritical, "Error"
    lblwait.Caption = ""
End Sub

Private Sub Form_Load()
    isDownload = False
    UpdateList = FixPath(App.Path) & "updates.dat" 'This will be the update list that downloads
    'Below just adds some line effects
    ln.X2 = Me.ScaleWidth
    Line1(0).X2 = Me.ScaleWidth: Line1(1).X2 = Me.ScaleWidth
    Line2(0).X2 = Me.ScaleWidth: Line2(1).X2 = Me.ScaleWidth
    'Open mirros list
    OpenMirrros FixPath(App.Path) & "mirrors.txt"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close Program and clean up
    sTmp = 0
    Set frmMain = Nothing
    Erase Mirros
    Erase DataList
    End Sub

Private Sub lstUpdates_Click()
Dim vLst As Variant, iSize As Integer, Idx As Integer
Dim sLine As String

    Idx = (lstUpdates.ListIndex + 1)
    sLine = DataList(Idx)
    vLst = Split(sLine, "|")
    sLine = ""
    
    If UBound(vLst) = 3 Then
        lblVer.Caption = vLst(1)
        lblSize.Caption = vLst(2)
        lblUrl.Caption = vLst(3)
    End If
    
    Erase vLst

    cmdDown.Enabled = True
    
End Sub
