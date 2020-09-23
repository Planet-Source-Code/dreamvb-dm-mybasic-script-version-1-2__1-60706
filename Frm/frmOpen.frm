VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Environment Options"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicTab1 
      BorderStyle     =   0  'None
      Height          =   3345
      Index           =   1
      Left            =   225
      ScaleHeight     =   3345
      ScaleWidth      =   6585
      TabIndex        =   4
      Top             =   615
      Visible         =   0   'False
      Width           =   6585
      Begin VB.Frame Frame2 
         Caption         =   "Editor"
         Height          =   3015
         Left            =   105
         TabIndex        =   5
         Top             =   315
         Width           =   6165
         Begin VB.ListBox lstStyles 
            Height          =   990
            IntegralHeight  =   0   'False
            Left            =   2580
            TabIndex        =   32
            Top             =   525
            Width           =   975
         End
         Begin VB.TextBox txtMargin 
            Height          =   300
            Left            =   5160
            TabIndex        =   26
            Text            =   "5"
            Top             =   660
            Width           =   690
         End
         Begin VB.TextBox txtTabSize 
            Height          =   300
            Left            =   5160
            TabIndex        =   24
            Text            =   "4"
            Top             =   232
            Width           =   690
         End
         Begin VB.Frame Frame3 
            Height          =   885
            Left            =   135
            TabIndex        =   10
            Top             =   1905
            Width           =   2235
            Begin VB.PictureBox PicEdBkCol 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   1125
               ScaleHeight     =   195
               ScaleWidth      =   765
               TabIndex        =   14
               Top             =   465
               Width           =   825
            End
            Begin VB.PictureBox PicEdFCol 
               BackColor       =   &H00000000&
               Height          =   240
               Left            =   105
               ScaleHeight     =   180
               ScaleWidth      =   765
               TabIndex        =   12
               Top             =   465
               Width           =   825
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Back Colour:"
               Height          =   195
               Left            =   1125
               TabIndex        =   13
               Top             =   195
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fore Colour:"
               Height          =   195
               Left            =   105
               TabIndex        =   11
               Top             =   195
               Width           =   855
            End
         End
         Begin VB.ComboBox cbofontsize 
            Height          =   315
            Left            =   195
            TabIndex        =   9
            Top             =   1185
            Width           =   2175
         End
         Begin VB.ComboBox cboFont 
            Height          =   315
            Left            =   195
            TabIndex        =   6
            Top             =   525
            Width           =   2175
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FontStyle:"
            Height          =   195
            Left            =   2610
            TabIndex        =   31
            Top             =   285
            Width           =   705
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left Margin Size"
            Height          =   195
            Left            =   3870
            TabIndex        =   25
            Top             =   735
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tab Size"
            Height          =   195
            Left            =   3870
            TabIndex        =   23
            Top             =   285
            Width           =   630
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Editor Colour Properties:"
            Height          =   195
            Left            =   210
            TabIndex        =   15
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label l10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font Size:"
            Height          =   195
            Left            =   210
            TabIndex        =   8
            Top             =   960
            Width           =   705
         End
         Begin VB.Label l6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font:"
            Height          =   195
            Left            =   210
            TabIndex        =   7
            Top             =   285
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox PicTab1 
      BorderStyle     =   0  'None
      Height          =   3885
      Index           =   0
      Left            =   150
      ScaleHeight     =   3885
      ScaleWidth      =   6585
      TabIndex        =   3
      Top             =   5715
      Visible         =   0   'False
      Width           =   6585
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Recent List"
         Height          =   375
         Left            =   4665
         TabIndex        =   33
         Top             =   1470
         Width           =   1740
      End
      Begin VB.TextBox txtHeader 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   165
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   2535
         Width           =   4305
      End
      Begin VB.TextBox txtMaxDocs 
         Height          =   315
         Left            =   3690
         TabIndex        =   28
         Top             =   1485
         Width           =   855
      End
      Begin VB.CheckBox chkMax 
         Caption         =   "Always-Open Environment Maxsized"
         Height          =   225
         Left            =   180
         TabIndex        =   22
         Top             =   1950
         Width           =   2940
      End
      Begin VB.CommandButton cmdMerge 
         Caption         =   "...."
         Height          =   315
         Left            =   5580
         TabIndex        =   21
         Top             =   1020
         Width           =   360
      End
      Begin VB.TextBox txtMerge 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1035
         Width           =   5355
      End
      Begin VB.CommandButton cmdEngine 
         Caption         =   "...."
         Height          =   315
         Left            =   5580
         TabIndex        =   18
         Top             =   315
         Width           =   360
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   330
         Width           =   5355
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Comment Header:"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   2295
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum number of recent documents to show."
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   1545
         Width           =   3375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "My Basic-Script Merge tool:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   795
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MyBasic Engine Path:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   30
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5745
      TabIndex        =   1
      Top             =   4815
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4395
      TabIndex        =   0
      Top             =   4815
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   5535
      Top             =   6660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip sTab 
      Height          =   4530
      Left            =   135
      TabIndex        =   2
      Top             =   150
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   7990
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "T_GEN"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp1 As String, Temp2 As String

Function SerachListBox(lpSerachFor As String, cboBox As ComboBox) As Integer
Dim n As Integer

    SerachListBox = -1
    For n = 0 To cboBox.ListCount
        If LCase$(cboBox.List(n)) = LCase$(lpSerachFor) Then
            SerachListBox = n
            Exit For
        End If
    Next n
    
End Function

Sub ReAlignTabs(nIndex As Integer)
    PicTab1(nIndex).Visible = True
    PicTab1(nIndex).Top = 615
    PicTab1(nIndex).Left = 255
End Sub

Sub DoColor(mPicBox As PictureBox)
On Error GoTo DlgError
    With CDLG
        .CancelError = True
        .ShowColor
        mPicBox.BackColor = .Color
        Exit Sub
DlgError:
        If Err = cdlCancel Then Err.Clear
    End With
    
End Sub

Sub FixSizes(bTextBox As TextBox, lpDefault As Integer)
    If Not IsNumeric(bTextBox.Text) Then bTextBox.Text = lpDefault: Exit Sub
    If Val(bTextBox.Text) < 0 Then bTextBox.Text = lpDefault
End Sub

Function SelectFile() As String
On Error GoTo DlgError
    With CDLG
        .CancelError = True
        .DialogTitle = "Open"
        .InitDir = MainAppPath & "engine\"
        .Filter = "Program Files(*.exe)|*.exe"
        .ShowOpen
        If Len(.Filename) = 0 Then Exit Function
        SelectFile = .Filename
        Exit Function
DlgError:
        If Err = cdlCancel Then Err.Clear
    End With
    
End Function

Private Sub cboFont_Change()
    cboFont.Text = Temp1
End Sub

Private Sub cboFont_Click()
    Temp1 = cboFont.Text
End Sub

Private Sub cbofontsize_Change()
    cbofontsize.Text = Temp2
End Sub

Private Sub cbofontsize_Click()
    Temp2 = cbofontsize.Text
End Sub

Private Sub cmdCancel_Click()
    Temp1 = "": Temp2 = ""
    Unload frmOptions
End Sub

Private Sub cmdClear_Click()
    If MsgBox("Are you sure you want to clear the entire recent list?", _
        vbYesNo Or vbQuestion, cmdClear.Caption) = vbNo Then
        Exit Sub
    Else
        SaveFile Recent_File, ""
        LoadHomePage frmMain.WebV
    End If
    
End Sub

Private Sub cmdEngine_Click()
Dim s As String
    s = SelectFile
    If Len(s) > 0 Then txtPath.Text = s
    s = ""
End Sub

Private Sub cmdMerge_Click()
Dim s As String
    s = SelectFile
    If Len(s) > 0 Then txtMerge.Text = s
    s = ""
End Sub

Private Sub cmdok_Click()
    'General selection
    t_Config.nEngine = txtPath.Text
    t_Config.nMergeTool = txtMerge.Text
    t_Config.nRecDocMax = Val(txtMaxDocs.Text)
    t_Config.nFullSizeWindow = chkMax.Value
    'Editor selection
    t_Config.nFont = Temp1
    t_Config.nFontSize = Val(Temp2)
    t_Config.nFontStyle = lstStyles.ListIndex
    t_Config.nBackColor = PicEdBkCol.BackColor
    t_Config.nForeColor = PicEdFCol.BackColor
    t_Config.nLeftMargin = Val(txtMargin.Text)
    t_Config.nTabSize = Val(txtTabSize.Text)
    SaveFile DataPath & "Comment header.txt", txtHeader.Text
    'Write the config settings and load the new ones backin
    WriteToConfig
    txtHeader.Text = OpenFile(DataPath & "Comment header.txt")
    LoadConfig
    cmdCancel_Click
End Sub

Private Sub Form_Load()
Dim X As Integer
    
    Set frmOptions.Icon = Nothing
    
    If IsFileHere(DataPath & "Comment header.txt") Then
        txtHeader.Text = OpenFile(DataPath & "Comment header.txt")
    End If
    
    For X = 0 To Screen.FontCount - 1
        cboFont.AddItem Screen.Fonts(X)
    Next
    
    cbofontsize.AddItem "8"
    cbofontsize.AddItem "9"
    cbofontsize.AddItem "10"
    cbofontsize.AddItem "12"
    cbofontsize.AddItem "14"
    cbofontsize.AddItem "16"
    cbofontsize.AddItem "18"
    cbofontsize.AddItem "20"
    cbofontsize.AddItem "22"
    cbofontsize.AddItem "24"

    lstStyles.AddItem "None"
    lstStyles.AddItem "Bold"
    lstStyles.AddItem "Italic"
    lstStyles.AddItem "Bold+Italic"
    
    sTab_Click
    
    'Load config data
    txtPath.Text = t_Config.nEngine
    txtMerge.Text = t_Config.nMergeTool
    txtMaxDocs.Text = t_Config.nRecDocMax
    chkMax.Value = t_Config.nFullSizeWindow
    txtTabSize.Text = t_Config.nTabSize
    txtMargin.Text = t_Config.nLeftMargin
    PicEdFCol.BackColor = t_Config.nForeColor
    PicEdBkCol.BackColor = t_Config.nBackColor
    
    X = SerachListBox(t_Config.nFont, cboFont)
    
    If X = -1 Then cboFont.ListIndex = 0 Else cboFont.ListIndex = X
    
    X = SerachListBox(CStr(t_Config.nFontSize), cbofontsize)
    
    If X = -1 Then cbofontsize.ListIndex = 0 Else cbofontsize.ListIndex = X
    
    X = 0
    
    lstStyles.ListIndex = t_Config.nFontStyle
    
End Sub

Private Sub PicEdBkCol_Click()
    DoColor PicEdBkCol
End Sub

Private Sub PicEdFCol_Click()
    DoColor PicEdFCol
End Sub

Private Sub sTab_Click()
    PicTab1(0).Visible = False
    PicTab1(1).Visible = False
    ReAlignTabs sTab.SelectedItem.Index - 1
End Sub

Private Sub txtMargin_LostFocus()
    FixSizes txtMargin, 5
End Sub

Private Sub txtMaxDocs_LostFocus()
    FixSizes txtMaxDocs, 8
End Sub

Private Sub txtTabSize_LostFocus()
    FixSizes txtTabSize, 4
End Sub
