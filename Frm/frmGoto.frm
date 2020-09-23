VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Goto"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Line3D1 
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   0
      Width           =   795
   End
   Begin VB.TextBox txtGoto 
      Height          =   330
      Left            =   2280
      TabIndex        =   5
      Top             =   525
      Width           =   3270
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Goto"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1215
      Width           =   1215
   End
   Begin VB.ListBox lstGoto 
      Height          =   1035
      Left            =   135
      TabIndex        =   1
      Top             =   525
      Width           =   2025
   End
   Begin VB.Label lbltitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2280
      TabIndex        =   2
      Top             =   225
      Width           =   3180
   End
   Begin VB.Label lblgoto 
      AutoSize        =   -1  'True
      Caption         =   "Goto:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   390
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    ButtonPressed = 0
    lstGoto.Clear
    txtGoto.Text = ""
    Unload frmGoto
End Sub

Private Sub cmdGoto_Click()
    If (TSelectionType = 0) Or (TSelectionType = 1) Then
        ButtonPressed = 1
        lstGoto.Clear
        Unload frmGoto
        Exit Sub
    End If
    
    If Not IsNumeric(txtGoto.Text) Then
        MsgBox "Invaild numeric value entered" & vbCrLf & "Please try agian.", vbExclamation, frmGoto.Caption
        Exit Sub
    Else
        mGoto = Val(txtGoto.Text)
        ButtonPressed = 1
        lstGoto.Clear
        txtGoto.Text = ""
        Unload frmGoto
    End If
    
End Sub

Private Sub Form_Load()
    lstGoto.AddItem "Code Start"
    lstGoto.AddItem "Code End"
    lstGoto.AddItem "Selection"
    lstGoto.AddItem "Line"
    lstGoto.ListIndex = 0
    Line3D1.Width = (frmGoto.Width - Line3D1.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGoto = Nothing
End Sub

Private Sub lstGoto_Click()
    TSelectionType = lstGoto.ListIndex
    
    txtGoto.BackColor = vbWhite
    
    Select Case TSelectionType
        Case 0
            txtGoto.Enabled = False
            txtGoto.BackColor = &H8000000F
            lbltitle.Caption = "Code Start"
        Case 1
            txtGoto.Enabled = False
            txtGoto.BackColor = &H8000000F
            lbltitle.Caption = "Code End"
        Case 2
            txtGoto.Enabled = True
            txtGoto.SetFocus
            lbltitle.Caption = "Selection number:"
        Case 3
            txtGoto.Enabled = True
            txtGoto.SetFocus
            lbltitle.Caption = "Line number:"
    End Select
End Sub
