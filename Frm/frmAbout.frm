VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5655
      TabIndex        =   4
      Top             =   1890
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   6825
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DM My Basic-Script Environment."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   930
         TabIndex        =   5
         Top             =   150
         Width           =   2940
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   105
         Picture         =   "frmAbout.frx":0000
         Top             =   135
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2003 - 2004 Ben Jones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   4095
         TabIndex        =   3
         Top             =   495
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "THIS PROGRAM IS FREEWARE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4095
         TabIndex        =   2
         Top             =   795
         Width           =   2520
      End
      Begin VB.Label lblver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM MyBasic-Script"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4095
         TabIndex        =   1
         Top             =   150
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload frmAbout
End Sub

Private Sub Form_Load()
    lblver.Caption = "Beta 1.3"
    frmAbout.Caption = "About " & frmMain.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub
