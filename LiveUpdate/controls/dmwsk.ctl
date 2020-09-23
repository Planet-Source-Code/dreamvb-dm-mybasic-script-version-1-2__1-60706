VERSION 5.00
Begin VB.UserControl DmDownload 
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleWidth      =   240
   ToolboxBitmap   =   "dmwsk.ctx":0000
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "dmwsk.ctx":0312
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "DmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Event DownloadProgress(mCurBytes As Long, mMaxBytes As Long)
Event DownloadComplete(mCurBytes As Long, mMaxBytes As Long, LocalFile As String)
Event LastError(StatusCode As Long, Status As String, LocalFile As String)
Event Status(StatusText As String)
Event StatusCode(Code As AsyncStatusCodeConstants)

Public StillBusy As Boolean

Private Sub SaveData(mData() As Byte, LocalFile As String)
Dim TFile As Long
    TFile = FreeFile
    
    Open LocalFile For Binary Access Write As #TFile
        Put #TFile, , mData()
    Close #TFile
    
End Sub
Public Sub DownloadFile(URL As String, LocalFile As String, Optional mType As AsyncTypeConstants = vbAsyncTypeByteArray)
On Error GoTo ConErr
    UserControl.AsyncRead URL, mType, LocalFile, vbAsyncReadForceUpdate
    Exit Sub
ConErr:
    RaiseEvent LastError(0, Err.Description, LocalFile)
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error Resume Next
    StillBusy = False
    With AsyncProp
        If .BytesMax <> 0 Then
            SaveData .Value, .PropertyName
            RaiseEvent DownloadComplete(.BytesRead, .BytesMax, .PropertyName)
        Else
            RaiseEvent LastError(.StatusCode, .Status, .PropertyName)
            Exit Sub
        End If
    End With
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
On Error Resume Next

    If AsyncProp.StatusCode = vbAsyncStatusCodeError Then
        StillBusy = False
        RaiseEvent LastError(AsyncProp.StatusCode, AsyncProp.Status, AsyncProp.PropertyName)
        RaiseEvent Status(AsyncProp.Status)
        RaiseEvent StatusCode(AsyncProp.StatusCode)
    Else
        RaiseEvent Status(AsyncProp.Status)
        RaiseEvent DownloadProgress(AsyncProp.BytesRead, AsyncProp.BytesMax)
        RaiseEvent StatusCode(AsyncProp.StatusCode)
        StillBusy = True
    End If
    
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 240, 240
End Sub
