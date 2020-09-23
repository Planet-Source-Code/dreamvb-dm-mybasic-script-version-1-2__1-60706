VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "DM MyBasic-Script"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8955
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8280
      Top             =   1875
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin BasicDevEnv.Bevel Bevel1 
      Height          =   3165
      Left            =   75
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1185
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   5583
      BevelStyle      =   0
      Begin VB.PictureBox PicTitleBar 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   30
         ScaleHeight     =   315
         ScaleWidth      =   2385
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   30
         Width           =   2385
         Begin VB.PictureBox PicBut 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   30
            Picture         =   "frmMain.frx":0442
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lbTaskTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Functions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   300
            TabIndex        =   13
            Top             =   30
            Width           =   1005
         End
      End
      Begin MSComctlLib.TreeView Tv1 
         Height          =   2760
         Left            =   30
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   4868
         _Version        =   393217
         Indentation     =   2
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   0
      End
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1185
      Width           =   1665
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1429
      BandCount       =   2
      _CBWidth        =   8955
      _CBHeight       =   810
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      NewRow1         =   0   'False
      MinHeight2      =   360
      UseCoolbarColors2=   0   'False
      NewRow2         =   -1  'True
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   990
         TabIndex        =   5
         Top             =   435
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
      End
      Begin VB.PictureBox picHolder 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   135
         ScaleHeight     =   315
         ScaleWidth      =   915
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   435
         Width           =   915
         Begin VB.Label lbladdins 
            AutoSize        =   -1  'True
            Caption         =   "Add-Ins"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   4
            Top             =   45
            Width           =   630
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   195
         TabIndex        =   2
         Top             =   45
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_NEW"
               Object.ToolTipText     =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_OPEN"
               Object.ToolTipText     =   "Open"
               ImageIndex      =   2
               Style           =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_SAVE"
               Object.ToolTipText     =   "Save"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_CUT"
               Object.Tag             =   "Cut"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_COPY"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_PASTE"
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_FIND"
               Object.ToolTipText     =   "Find / Replace"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_RUN"
               Object.ToolTipText     =   "Run"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "M_STOP"
               Object.ToolTipText     =   "Stop"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "M_BUILD"
               Object.ToolTipText     =   "Build"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_START"
               Object.ToolTipText     =   "StartPage"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_UPDATE"
               Object.ToolTipText     =   "Check For Updates"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "M_HELP"
               Object.ToolTipText     =   "Help"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.ListBox LBError 
      Height          =   720
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":09CC
      Left            =   60
      List            =   "frmMain.frx":09CE
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   1920
   End
   Begin SHDocVwCtl.WebBrowser WebV 
      Height          =   645
      Left            =   75
      TabIndex        =   8
      Top             =   1185
      Visible         =   0   'False
      Width           =   1110
      ExtentX         =   1958
      ExtentY         =   1138
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4365
      Left            =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   825
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   7699
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Start Page"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6045
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9472
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   7860
      Top             =   1005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1074
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1718
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":210E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2460
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":384C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopen 
         Caption         =   "O&pen..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnublank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnublank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuclear 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnublank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find / Replace..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnublank7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselall 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuBlnak8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconv 
         Caption         =   "C&onvert"
         Begin VB.Menu mnucase1 
            Caption         =   "&UpperCase"
         End
         Begin VB.Menu mnucase2 
            Caption         =   "&LowerCase"
         End
         Begin VB.Menu mnuTab 
            Caption         =   "&Tabs to Spaces"
         End
      End
      Begin VB.Menu mnuinsert 
         Caption         =   "&Insert"
         Begin VB.Menu mnudate 
            Caption         =   "Date"
         End
         Begin VB.Menu mnutime 
            Caption         =   "Time"
         End
         Begin VB.Menu mnudatetime 
            Caption         =   "Date/Time"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnublank9 
            Caption         =   "-"
         End
         Begin VB.Menu mnucomhead 
            Caption         =   "&Comment &Header"
         End
         Begin VB.Menu mnuInstFile 
            Caption         =   "File..."
         End
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuoptions 
         Caption         =   "Options..."
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "&Check for Updates"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelp2 
         Caption         =   "Engine Help"
         Begin VB.Menu mnuindex 
            Caption         =   "Index..."
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuchnage 
            Caption         =   "Change Log"
         End
         Begin VB.Menu mnublank3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuvar 
            Caption         =   "Variables"
         End
         Begin VB.Menu mnuconst 
            Caption         =   "&Const"
         End
         Begin VB.Menu mnudata 
            Caption         =   "Data Types"
         End
         Begin VB.Menu mnublank6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRef 
            Caption         =   "Functions Reference"
         End
         Begin VB.Menu mnykeyref 
            Caption         =   "Keywords Reference"
         End
      End
      Begin VB.Menu mnubalnk9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuchLog 
         Caption         =   "&What's New"
      End
      Begin VB.Menu mnublank10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Changed As Boolean
Dim Tab_Show As Integer
Dim Ide_First_Load As Boolean
Dim Doc_File As String

Const Env_Reg_Path = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"

Sub AddinToolBarEnable(mEnable As Boolean)
Dim I As Integer
    'Enables or disables the add-ins toolbar
    For I = 1 To Toolbar2.Buttons.Count
        Toolbar2.Buttons(1).Enabled = mEnable
    Next
    
End Sub

Private Sub InstallPlugins()
Dim X As String, plgFile As String, plgServ As String
Dim Idx As Integer
Dim PlgObj As Object
On Error Resume Next

    If Dir(Plg_Path) = "" Then Exit Sub
    
    X = Dir(Plg_Path)
    ImageList2.ListImages.Clear
    Idx = 0
    
    Do While X <> ""
        If GetFileExt(X) = "DLL" Then
            plgFile = Plg_Path & X
            plgServ = GetFileTitle(X) & "main"
            If RegisterActiveX(plgFile, Register) = False Then
                MsgBox "There was an error loading the add-in:" _
                & vbCrLf & Plg_Path & X
            Else
                Set PlgObj = CreateObject(plgServ)
                If PlgObj.loadplug <> 1 Then
                    MsgBox "There was an error loading the add-in:", vbCritical
                Else
                    Idx = Idx + 1
                    ImageList2.ListImages.Add , , PlgObj.ipic
                    Toolbar2.Buttons.Add , plgServ, , , Idx
                    Toolbar2.Buttons(Idx).ToolTipText = PlgObj.ToolTipText
                    Set PlgObj = Nothing
                    plgServ = ""
                    plgFile = ""
                End If
            End If
        End If
        X = Dir
        DoEvents
    Loop
    
    AddinToolBarEnable False
End Sub

Private Sub SetSysEnvPath()
Dim sEnvPath As String
Dim sPath As String

    sPath = GetShPath(MainAppPath) 'Path to this program
    sPath = Left(sPath, Len(sPath) - 1) 'Remove the last backslash
    sEnvPath = RegReadString(HKEY_LOCAL_MACHINE, Env_Reg_Path, "DM_MyBasic", REG_EXPAND_SZ)
    
    If Len(sEnvPath) = 0 Then
        'Set new system env key
        RegSaveValue HKEY_LOCAL_MACHINE, Env_Reg_Path, "DM_MyBasic", REG_EXPAND_SZ, sPath
        'We need to refresh the enviroment so we call our little C++ program
        If Not IsFileHere(MainAppPath & "EnvR.exe") Then Exit Sub
        RunFile MainAppPath & "EnvR.exe", hwnd, "", 0
    Else
        'Now we check sPath is the same as sEnvPath
        ' if we don;t get a match then the program has been moved and we need to reset it
        If Not (sEnvPath = sPath) Then
            RegSaveValue HKEY_LOCAL_MACHINE, Env_Reg_Path, "DM_MyBasic", REG_EXPAND_SZ, sPath
            If Not IsFileHere(MainAppPath & "EnvR.exe") Then Exit Sub
            RunFile MainAppPath & "EnvR.exe", hwnd, "", 0
        End If
        'Clear up
        Env_Path = "": sPath = ""
    End If
End Sub

Sub OpenScript(bFile As String)
    If Not IsFileHere(bFile) Then
        MsgBox "File not found:" & vbCrLf & bFile, vbCritical, "File not found"
        Exit Sub
    Else
        txtCode.Text = OpenFile(bFile)
        Changed = False
        StatusBar1.Panels(1).Text = "" ' clear the statbar panel
        lzScript_File = bFile
        Ide_First_Load = False
    End If
    
End Sub

Sub DoWebOpen(lzFile As String)
Dim Ans As Integer

    If Not Changed Then
        OpenScript lzFile
        AdEditorTab
        TabStrip1.Tabs(2).Selected = True
        Exit Sub
    End If
    
    Ans = MsgBox("You have made chnages to your work." _
    & vbCrLf & vbCrLf & "Do you want to save the chnages now?", vbYesNo Or vbQuestion, frmMain.Caption)
        
    If Ans = vbNo Then
        OpenScript lzFile
        AdEditorTab
        TabStrip1.Tabs(2).Selected = True
        Exit Sub
    Else
        DoSave
        OpenScript lzFile
        AdEditorTab
        TabStrip1.Tabs(2).Selected = True
    End If

End Sub

Public Sub AddError(lpError As String)
Dim vLst As Variant, X As Integer
    
    Beep 'Sound an error
    Toolbar1.Buttons(11).Enabled = False
    
    If Len(lpError) <> 0 Then
        vLst = Split(lpError, vbCrLf)
        For X = 0 To UBound(vLst)
            LBError.AddItem vLst(X)
        Next
    End If
    
    LBError.ListIndex = (LBError.ListCount - 1)
    X = 0
    Erase vLst
    
End Sub

Sub AdEditorTab()
    If Ide_First_Load Then Exit Sub
    If TabStrip1.Tabs.Count > 1 Then Exit Sub 'Tab alreadi exsits no need to add new one
    TabStrip1.Tabs.Add , "code", "Editor" 'Add a new tab for the editor
    TabStrip1.Tabs(2).Selected = True 'Set focus on the editor tab
End Sub

Sub IdeShopwTab(Index As Integer)
    Tab_Show = Index
    
    If Index = 1 Then
        Bevel1.Visible = False
        txtCode.Visible = False
        WebV.Visible = True
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Toolbar1.Buttons(10).Enabled = False
        Toolbar1.Buttons(8).ToolTipText = "Find"
        mnuedit.Enabled = False
        mnuRun.Enabled = False
        mnusave.Enabled = False
        LBError.Visible = False
        StatusBar1.Panels(2).Visible = False: StatusBar1.Panels(3).Visible = False: _
        StatusBar1.Panels(4).Visible = False: StatusBar1.Panels(5).Visible = False
        AddinToolBarEnable False
    Else
        AddinToolBarEnable True
        StatusBar1.Panels(2).Visible = True: StatusBar1.Panels(3).Visible = True: _
        StatusBar1.Panels(4).Visible = True: StatusBar1.Panels(5).Visible = True
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(10).Enabled = True
        Toolbar1.Buttons(8).ToolTipText = "Find / Replace"
        Bevel1.Visible = True
        LBError.Visible = True
        txtCode.Visible = True
        WebV.Visible = False
        mnuedit.Enabled = True
        mnuRun.Enabled = True
        mnusave.Enabled = True
        EnableEditMenu
        txtCode.SetFocus
        If Tv1.Nodes.Count = 0 Then
            frmMenu2.mnukeys_Click
            Tv1.Nodes(2).Selected = True
        End If
        
    End If
    
    Form_Resize
End Sub

Private Sub DoOpen()
On Error GoTo DlgError
    With CDialog
        .CancelError = True ' Turn on error checking
        .DialogTitle = "Open" ' dialog title
        .Filter = dlg_filter ' Filter file type
        .InitDir = MainAppPath & "examples"
        .ShowOpen ' show the save dialog
        If Len(.Filename) = 0 Then Exit Sub: Ide_First_Load = True
        'Add file to recenet
        If IsInList(.Filename) <> True Then
            WriteRecentList .Filename
        End If
        LoadHomePage WebV
        txtCode.Text = OpenFile(.Filename)
        Changed = False
        StatusBar1.Panels(1).Text = "" ' clear the statbar panel
        lzScript_File = .Filename
        Ide_First_Load = False
        Exit Sub ' exit code block
DlgError:
        If Err = cdlCancel Then Err.Clear
        Ide_First_Load = True
    End With

End Sub

Private Sub DoSave()
On Error GoTo DlgError
    With CDialog
        .CancelError = True ' Turn on error checking
        .DialogTitle = "Save" ' dialog title
        .Filter = dlg_filter ' Filter file type
        .ShowSave ' show the save dialog
        .InitDir = MainAppPath & "examples"
         SaveFile .Filename, txtCode.Text ' Save the data in the editor
         If Len(.Filename) = 0 Then Exit Sub: Ide_First_Load = True
         lzScript_File = .Filename
         Changed = False
         Ide_First_Load = False
         Exit Sub ' exit code block
DlgError:
        If Err = cdlCancel Then Err.Clear
        Ide_First_Load = True
    End With

End Sub

Private Sub EnableEditMenu()
    'Menu items
    mnucut.Enabled = clsTextBox.EnableCutPaste
    mnucopy.Enabled = clsTextBox.EnableCutPaste
    mnupaste.Enabled = Not clsTextBox.IsClipEmpty
    mnuclear.Enabled = mnucut.Enabled
    mnuconv.Enabled = mnucut.Enabled
    'Toolbar buttons
    Toolbar1.Buttons(5).Enabled = mnucut.Enabled
    Toolbar1.Buttons(6).Enabled = mnucopy.Enabled
    Toolbar1.Buttons(7).Enabled = mnupaste.Enabled
End Sub


Private Sub Form_Load()
Dim Start_Path As String
    
    MainAppPath = FixPath(App.Path)
    Plg_Path = MainAppPath & "add-ins\"
    SetSysEnvPath 'Add a system evnviroment string for the path of this program
    Config_File = MainAppPath & App.EXEName & ".ini"
    Doc_File = MainAppPath & "docs\help.chm"
    
    DataPath = MainAppPath & "data\" 'DataPath
    Recent_File = DataPath & "recent.ini"
    Start_Path = DataPath & "startpage\"
    StartPage = Start_Path & "index.htm"
    WebTemp = Start_Path & "web.tmp"
    IncFile = Start_Path & "Recent.inc"
    
    'Crate a new Recent file unless one already exsits
    If Not IsFileHere(Recent_File) Then SaveFile DataPath & "Recent.ini", ""
    
    If Not IsFileHere(IncFile) Then
        If Not IsFileHere(StartPage) Then
            WebV.Navigate "about:blank"
        Else
            WebTemp = StartPage
            WebV.Navigate StartPage
        End If
    Else
        If Not IsFileHere(StartPage) Then WebV.Navigate "about:blank" Else LoadHomePage WebV
    End If
    
    clsTextBox.TextBox = txtCode
    'Load in editor configuration
    IniClass.iniFile = Config_File
    'If no ini we must write the defaults
    If Not IsFileHere(Config_File) Then ConfigDefaults
    'Load in config
    LoadConfig
    
    If t_Config.nFullSizeWindow = 1 Then frmMain.WindowState = 2
    EnableEditMenu ' enable the edit menu
    TabStrip1_Click
    
    mWnd = frmMain.hwnd 'Get the forms Hangle
    Hook  'Place a hook on this form
    Start_Path = ""
    InstallPlugins
    DoEvents
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mnuexit_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    If frmMain.Width < 3930 Then frmMain.Width = 3930
    If frmMain.Height < 2460 Then frmMain.Height = 2460
    
    TabStrip1.Width = frmMain.ScaleWidth
    TabStrip1.Height = (frmMain.ScaleHeight - TabStrip1.Top - StatusBar1.Height)
    
    If Tab_Show = 2 Then
        Bevel1.Height = (TabStrip1.Height - Bevel1.Top)
        Tv1.Height = (Bevel1.Height - Tv1.Top - 20)
        txtCode.Left = Bevel1.Width + 60
        
        txtCode.Width = (TabStrip1.Width - txtCode.Left - 40)
        LBError.Width = (TabStrip1.Width - LBError.Left - 40)
        LBError.Top = TabStrip1.Height
        txtCode.Height = (LBError.Top - txtCode.Top - 40)
    Else
        WebV.Width = (TabStrip1.Width - WebV.Left - 40)
        WebV.Height = (TabStrip1.Height - StatusBar1.Height - 140)
    End If
    
    DoEvents
    
End Sub

Private Sub Form_Terminate()
    If IsFileHere(WebTemp) Then Kill WebTemp
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set frmMain = Nothing
    Set frmFindReplace = Nothing
    Set frmGoto = Nothing
    Set frmMenu2 = Nothing
    Set frmOptions = Nothing
    Call UnHook
    If IsFileHere(WebTemp) Then Kill WebTemp

End Sub

Private Sub LBError_Click()
Dim e_pos As Integer, sTmp As String, Ln_Idx As Long
    sTmp = LBError.Text 'Get the text from the selected item in the listbox
    e_pos = InStr(1, sTmp, ":", vbBinaryCompare) 'Start position to look for
    If e_pos <> 0 Then
        'Check for the line error
        If UCase(Left(sTmp, e_pos - 1)) = "LINE" Then
            'Extract the line number
            Ln_Idx = CLng(Trim(Mid(sTmp, e_pos + 1, Len(sTmp))))
            If Ln_Idx <> 0 Then
                'If line number is greator then zero move to that line in the editor
                clsTextBox.HighLightLine Ln_Idx - 1
            End If
        End If
    End If
    
End Sub

Private Sub mnuabout_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnucase1_Click()
    txtCode.SelText = UCase$(txtCode.SelText)
End Sub

Private Sub mnucase2_Click()
    txtCode.SelText = LCase$(txtCode.SelText)
End Sub

Private Sub mnuchLog_Click()
    WebV.Navigate DataPath & "startpage\whatsnew.htm"
End Sub

Private Sub mnuchnage_Click()
    WebV.Navigate "mk:@MSITStore:" & Doc_File & "::/changelog.htm#whatsnew"
End Sub

Private Sub mnuclear_Click()
    clsTextBox.EditMenu vsDELETE
    EnableEditMenu
End Sub

Private Sub mnucomhead_Click()
Dim Old_Pos As Long, sText As String

    If IsFileHere(DataPath & "Comment header.txt") Then
        Old_Pos = txtCode.SelStart
        sText = OpenFile(DataPath & "Comment header.txt")
        sText = Replace(sText, "%User", Environ("USERNAME"), , , vbTextCompare)
        sText = Replace(sText, "%Date", Date, , , vbTextCompare)
        sText = Replace(sText, "%Time", frmMain.Caption, , , vbTextCompare)
        sText = Replace(sText, "%EnvName", frmMain.Caption, , , vbTextCompare)
        sText = Replace(sText, "%OS", Environ("OS"), , , vbTextCompare)
        
        If Len(lzScript_File) <> 0 Then
            sText = Replace(sText, "%FileName", GetFileName(lzScript_File), , , vbTextCompare)
            sText = Replace(sText, "%FileSize", FileLen(lzScript_File) & " bytes", , , vbTextCompare)
            sText = Replace(sText, "%ScriptPath", CurDir(lzScript_File), , , vbTextCompare)
        End If
        
        txtCode.SelStart = 0
        txtCode.SelText = sText
        txtCode.SelStart = Old_Pos
        sText = ""
    End If
    
End Sub

Private Sub mnuconst_Click()
    WebV.Navigate "mk:@MSITStore:" & Doc_File & "::/Consts.htm"
End Sub

Private Sub mnucopy_Click()
    clsTextBox.EditMenu vsCOPY
    EnableEditMenu
End Sub

Private Sub mnucut_Click()
    clsTextBox.EditMenu vsCUT
    EnableEditMenu
End Sub

Private Sub mnudata_Click()
    WebV.Navigate "mk:@MSITStore:" & Doc_File & "::/datatypes.htm"
End Sub

Private Sub mnudate_Click()
    txtCode.SelText = Date
End Sub

Private Sub mnudatetime_Click()
    txtCode.SelText = Now
End Sub

Private Sub mnuexit_Click()
Dim Ans As Integer

    If Not Changed Then Unload frmMain: Exit Sub
    
    Ans = MsgBox("You have made changes to your work." _
    & vbCrLf & vbCrLf & "Do you want to save the changes now?", vbYesNo Or vbQuestion, frmMain.Caption)
    
    If Ans = vbNo Then
        Changed = False
        Unload frmMain
        Exit Sub
    Else
        DoSave
        Unload frmMain
    End If
End Sub

Private Sub mnufind_Click()
    If Tab_Show = 1 Then
        'Web Broswer is showing
        WebV.SetFocus
        SendKeys "^f"
    Else
        frmFindReplace.Show , frmMain
    End If
    
End Sub

Private Sub mnuGoto_Click()
    frmGoto.Show vbModal, frmMain
    If ButtonPressed = 0 Then Exit Sub ' cancel button Pressed
    
    Select Case TSelectionType
        Case 0 ' goto start of code top line
            txtCode.SelStart = 0
            txtCode.SetFocus
        Case 1 ' goto bottom of code last line
            txtCode.SelStart = Len(txtCode.Text)
            txtCode.SetFocus
        Case 2 ' goto a selection in the code
            txtCode.SelStart = txtCode.SelStart + mGoto
        Case 3
            clsTextBox.GotoLine = CLng(mGoto - 1)
    End Select
End Sub

Private Sub mnuindex_Click()
    RunFile Doc_File, frmMain.hwnd, "", 1
End Sub

Private Sub mnuInstFile_Click()
On Error GoTo DlgError
    With CDialog
        .CancelError = True
        .DialogTitle = "Insert File"
        .Filter = "All Files(*.*)|*.*"
        .ShowOpen
        If Len(.Filename) = 0 Then Exit Sub
        txtCode.SelText = OpenFile(.Filename)
        Exit Sub
DlgError:
        If Err = cdlCancel Then Err.Clear
    End With
    
End Sub

Private Sub mnunew_Click()
Dim Ans As Integer
    
    Ide_First_Load = False
    
    If Not Changed Then
        txtCode.Text = ""
        Changed = False
        StatusBar1.Panels(1).Text = ""
        lzScript_File = ""
        AdEditorTab
        Exit Sub
    End If
    
    Ans = MsgBox("You have made changes to your work." _
    & vbCrLf & vbCrLf & "Do you want to save the changes now?", vbYesNo Or vbQuestion, frmMain.Caption)
    
    If Ans = vbNo Then
        txtCode.Text = ""
        Changed = False
        StatusBar1.Panels(1).Text = ""
        lzScript_File = ""
        AdEditorTab
        Exit Sub
    Else
        DoSave
        lzScript_File = ""
    End If
End Sub

Private Sub mnuopen_Click()
Dim Ans As Integer

    If Not Changed Then
        DoOpen
        AdEditorTab
        Exit Sub
    End If
    
    Ans = MsgBox("You have made chnages to your work." _
    & vbCrLf & vbCrLf & "Do you want to save the chnages now?", vbYesNo Or vbQuestion, frmMain.Caption)
        
    If Ans = vbNo Then
        DoOpen
        AdEditorTab
        Exit Sub
    Else
        DoSave
        DoOpen
        AdEditorTab
    End If
    
End Sub

Private Sub mnuoptions_Click()
    frmOptions.Show vbModal, frmMain
End Sub

Private Sub mnupaste_Click()
    clsTextBox.EditMenu vsPASTE
End Sub

Private Sub mnuRef_Click()
    WebV.Navigate "mk:@MSITStore:" & Doc_File & "::/Functions/contents.htm"
End Sub

Private Sub mnusave_Click()
    DoSave
End Sub

Private Sub mnuselall_Click()
    clsTextBox.EditMenu vsSELALL
    EnableEditMenu
End Sub

Private Sub mnuStart_Click()
    If GetConsolWnd <> 0 Then Exit Sub
    Toolbar1.Buttons(12).Enabled = False 'Disable build button
    LBError.Clear 'Clear error lisbox
    
    'Make sure engine file is found
    If Not IsFileHere(lzEngine_File) Then
        MsgBox "MyBasicScript Engine Not Found:" & vbCrLf & lzEngine_File, vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    'Check that the user has saved the script-file
    If Len(lzScript_File) = 0 Then
        MsgBox "Please save your script before running.", vbInformation
        Exit Sub
    End If
    
    Toolbar1.Buttons(11).Enabled = True 'Enable the stop button
    SaveFile GetShPath(lzScript_File), txtCode.Text 'Save the script file
    RunFile lzEngine_File, Me.hwnd, GetShPath(lzScript_File), 1 'Run the script

End Sub

Private Sub mnuTab_Click()
    txtCode.SelText = Replace(txtCode.SelText, vbTab, Chr(32))
End Sub

Private Sub mnutime_Click()
    txtCode.SelText = Time
End Sub

Private Sub mnuupdate_Click()
Dim UpdateFile As String
Dim mUpdateObj As Object
On Error Resume Next

    UpdateFile = MainAppPath & "LiveUpdate\LiveUpdate.dll"
    If Not IsFileHere(UpdateFile) Then
        MsgBox "Unable locate update program.", vbExclamation, "File Not Found"
        Exit Sub
    Else
        Set mUpdateObj = CreateObject("dmLiveUpdate.ClsLiveUpdate")
        mUpdateObj.LoadUpdate
    End If
    
    If Err Then MsgBox Err.Description, vbCritical
    
End Sub

Private Sub mnuvar_Click()
    WebV.Navigate "mk:@MSITStore:" & Doc_File & "::/Variables.htm"
End Sub

Private Sub mnykeyref_Click()
    WebV.Navigate "mk:@MSITStore:" & Doc_File & "::/keywords/contents.htm"
End Sub

Private Sub PicBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then PopupMenu frmMenu2.mnuMenu, , X
End Sub

Private Sub TabStrip1_Click()
    IdeShopwTab TabStrip1.SelectedItem.Index
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "M_NEW"
            mnunew_Click
        Case "M_OPEN"
            mnuopen_Click
        Case "M_SAVE"
            mnusave_Click
        Case "M_CUT"
            mnucut_Click
        Case "M_COPY"
            mnucopy_Click
        Case "M_PASTE"
            mnupaste_Click
        Case "M_FIND"
            mnufind_Click
        Case "M_RUN"
            mnuStart_Click
        Case "M_START"
            TabStrip1.Tabs(1).Selected = True
            WebV.Navigate WebTemp
        Case "M_UPDATE"
            mnuupdate_Click
        Case "M_HELP"
            mnuindex_Click
        Case "M_BUILD"
            If Not IsFileHere(lzBuild_Tool) Then
                MsgBox "File not found:" & vbCrLf & lzBuild_Tool, vbCritical, "File Not Found"
                Exit Sub
            Else
                'Merge the script file. by using out merging tool
                RunFile lzBuild_Tool, frmMain.hwnd, lzScript_File, 1
                '
                MsgBox "The file has been compiled to:" & vbCrLf & vbCrLf _
                & GetFileTitle(lzScript_File) & "exe", vbInformation, "Build"
            End If
        Case "M_STOP"
            'Close the console window
            If GetConsolWnd <> 0 Then CloseConsoleWnd GetConsolWnd
            Toolbar1.Buttons(11).Enabled = False 'Disable the stop button
    End Select
    
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    DoWebOpen ButtonMenu.Text
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Plg_Obj As Object

    Set Plg_Obj = CreateObject(Button.Key)
    If Plg_Obj.loadplug <> 1 Then
        MsgBox "There was an error while loading the add-in.", vbCritical, "Error"
        Exit Sub
    Else
        Plg_Obj.InitAddin frmMain
    End If
    

End Sub

Private Sub Tv1_DblClick()
    If Tv1.Nodes.Count = 0 Then Exit Sub
    If Left(Tv1.Nodes(Tv1.SelectedItem.Index).Key, 2) <> "a:" Then
        'Insert the function name
        txtCode.SelText = Tv1.SelectedItem.Text
        txtCode.SetFocus
    End If
    
End Sub

Private Sub txtCode_Change()
    StatusBar1.Panels(2).Text = "Ln " & clsTextBox.GetCurrentLineNumber + 1 & ", " & "Col " & clsTextBox.GetColumn
    
    If Changed Then Exit Sub
    Changed = True
    StatusBar1.Panels(1).Text = "Modified"
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then txtCode.SelText = Space(t_Config.nTabSize): KeyAscii = 0
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then txtCode.SelText = "": EnableEditMenu
End Sub

Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(2).Text = "Ln " & clsTextBox.GetCurrentLineNumber + 1 & ", " & "Col " & clsTextBox.GetColumn
End Sub

Private Sub txtCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EnableEditMenu
End Sub

Private Sub WebV_TitleChange(ByVal Text As String)
Dim e_pos As Integer, n_pos As Integer, HelpId As String, sTmp As String

    e_pos = InStr(1, Text, "ref:", vbTextCompare)
    If e_pos <> 0 Then
        HelpId = Mid(Text, e_pos + 4, Len(Text))
        HelpId = Replace(HelpId, "?", "#")
        n_pos = InStr(1, HelpId, "(", vbBinaryCompare)
        
        If n_pos = 0 Then
            Select Case UCase(HelpId)
                Case "OPENFILE": mnuopen_Click
                Case "NEW": mnunew_Click
                Case "EXIT": mnuexit_Click: Exit Sub
            End Select
        Else
            sTmp = Trim(Mid(HelpId, n_pos + 1, Len(HelpId) - n_pos - 1))
            sTmp = Replace(sTmp, "%20", Chr(32))
            
            Select Case UCase(Left(HelpId, n_pos - 1))
                Case "OPEN": DoWebOpen sTmp
                Case "REFRESH": LoadHomePage WebV
                Case "REMOVE"
                    If MsgBox("Are you sure you want to remove this item" _
                        & "from the list?", vbYesNo Or vbQuestion) = vbYes Then
                        RemoveRecentFile sTmp
                        LoadHomePage WebV
                    End If
            End Select
        End If
        WebV.Navigate WebTemp
    End If
    
End Sub
