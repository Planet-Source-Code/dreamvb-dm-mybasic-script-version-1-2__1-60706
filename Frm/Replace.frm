VERSION 5.00
Begin VB.Form frmFindReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find / Replace"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtreplace 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   8
      Top             =   1230
      Width           =   3015
   End
   Begin VB.CommandButton cmdRepAll 
      Caption         =   "&Replace All"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3570
      TabIndex        =   6
      Top             =   1110
      Width           =   1095
   End
   Begin VB.CommandButton cmdRep 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3570
      TabIndex        =   5
      Top             =   630
      Width           =   1095
   End
   Begin VB.TextBox txtfind 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   0
      Top             =   435
      Width           =   3015
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "&Find"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3570
      TabIndex        =   1
      Top             =   150
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3570
      TabIndex        =   2
      Top             =   1620
      Width           =   1095
   End
   Begin VB.CheckBox chkmatch 
      Caption         =   "Match Case"
      Height          =   195
      Left            =   195
      TabIndex        =   3
      Top             =   1770
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace With:"
      Height          =   195
      Left            =   195
      TabIndex        =   7
      Top             =   945
      Width           =   1020
   End
   Begin VB.Label lblfind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find Text:"
      Height          =   195
      Left            =   195
      TabIndex        =   4
      Top             =   225
      Width           =   705
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Find text dilaog added by ben jones
Dim Pos As Integer

Private Sub cmdCancel_Click()
    txtfind.Text = ""
    txtreplace.Text = ""
    Unload frmFindReplace
End Sub

Private Sub cmdfind_Click()
Dim Compare As Integer

    If chkmatch Then Compare = 0 Else Compare = 1
    
    Pos = InStr(Pos + 1, clsTextBox.Text, txtfind.Text, Compare)
    
    If Pos > 0 Then
        clsTextBox.SelStart = (Pos - 1)
        clsTextBox.SelLength = Len(txtfind.Text)
        clsTextBox.SetFocus
        cmdfind.Caption = "Find Next"
    Else
        cmdfind.Caption = "Find"
        MsgBox "String " & txtfind.Text & " was not found", vbInformation, frmFindReplace.Caption
    End If
    
    Compare = 0
    ipos = 0
End Sub

Private Sub cmdRep_Click()
Dim Compare As Integer

    Pos = 1
    
    If chkmatch Then Compare = 0 Else Compare = 1
    
    Pos = InStr(Pos + 1, clsTextBox.Text, txtfind.Text, Compare)
    
    If Pos > 0 Then
        clsTextBox.SelStart = (Pos - 1)
        clsTextBox.SelLength = Len(txtfind.Text)
        clsTextBox.SelText = txtreplace.Text
        clsTextBox.SetFocus
    Else
        MsgBox "String " & txtfind.Text & " was not found", vbInformation, frmFindReplace.Caption
    End If
    
    Compare = 0
    ipos = 0
    
End Sub

Private Sub cmdRepAll_Click()
Dim Compare As Integer
    
Do
    Pos = 1
    
    If chkmatch Then Compare = 0 Else Compare = 1
    
    Pos = InStr(Pos + 1, clsTextBox.Text, txtfind.Text, Compare)
    
    If Pos > 0 Then
        clsTextBox.SelStart = (Pos - 1)
        clsTextBox.SelLength = Len(txtfind.Text)
        clsTextBox.SelText = txtreplace.Text
        clsTextBox.SetFocus
    End If
    
    Compare = 0
    ipos = 0
    DoEvents
    Loop Until Pos = 0
    
End Sub

Private Sub Form_Load()
    frmFindReplace.Icon = Nothing
    Pos = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFind = Nothing
End Sub

Private Sub txtfind_Change()
    cmdfind.Enabled = Len(txtfind.Text) <> 0
    cmdRep.Enabled = cmdfind.Enabled
    cmdRepAll.Enabled = cmdfind.Enabled
End Sub

