VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Add Testing"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2835
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get IDE Caption"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2250
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set IDE Text"
      Height          =   375
      Left            =   135
      TabIndex        =   2
      Top             =   2835
      Width           =   1380
   End
   Begin VB.TextBox txtTest 
      Height          =   2040
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   5115
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Get IDE Text"
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   2250
      Width           =   1380
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ide_Object As Object
'This is a direct link right to the IDE
'This means you have full access to the editor and IDE if needed
' See below for a basic examples

Private Sub cmd1_Click()
    'Gets text from the code editor
    txtTest.Text = Ide_Object.txtcode.Text
End Sub

Private Sub Command1_Click()
    'Sets new text to the editor
    Ide_Object.txtcode.Text = "Testing ......"
End Sub

Private Sub Command2_Click()
    'Get the IDE's window caption
    txtTest.Text = "IDE Caption is " & Ide_Object.Caption
End Sub

Private Sub Command3_Click()
    Unload frmMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub
