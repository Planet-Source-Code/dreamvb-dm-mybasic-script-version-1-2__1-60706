VERSION 5.00
Begin VB.Form frmMenu2 
   Caption         =   "Form1"
   ClientHeight    =   120
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   120
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMenu 
      Caption         =   "&Mnu"
      Begin VB.Menu mnufunc 
         Caption         =   "&Functions"
      End
      Begin VB.Menu mnukeys 
         Caption         =   "&Keywords"
      End
      Begin VB.Menu mnusysVar 
         Caption         =   "Others"
      End
   End
End
Attribute VB_Name = "frmMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnufunc_Click()
    'Load inbuilt functions list
    frmMain.lbTaskTitle.Caption = "Functions"
    LoadFunctionsLst frmMain.Tv1, DataPath & "functions.htd"
End Sub

Public Sub mnukeys_Click()
    'Load Keyword list
    frmMain.lbTaskTitle.Caption = "Functions"
    LoadFunctionsLst frmMain.Tv1, DataPath & "keywords.htd"
End Sub

Private Sub mnusysVar_Click()
    'Load system variables list
    frmMain.lbTaskTitle.Caption = "Functions"
    LoadFunctionsLst frmMain.Tv1, DataPath & "others.htd"
End Sub
