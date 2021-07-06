VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Lock Status"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   3480
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkStartWithWin 
      Caption         =   "Start with Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CheckBox chkHideOnStart 
      Caption         =   "Do not show this window on start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CheckBox chkTrayVisible 
      Caption         =   "Visible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Tag             =   "scroll"
      Top             =   1680
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkTrayVisible 
      Caption         =   "Visible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Tag             =   "caps"
      Top             =   1320
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkTrayVisible 
      Caption         =   "Visible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Tag             =   "num"
      Top             =   960
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.PictureBox pictKeyLock 
      Height          =   375
      Index           =   2
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      ToolTipText     =   "ScrollLock Status"
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox pictKeyLock 
      Height          =   375
      Index           =   1
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      ToolTipText     =   "CapsLock Status"
      Top             =   1320
      Width           =   375
   End
   Begin VB.PictureBox pictTrayIcon 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2640
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer tmrInitApp 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2640
      Top             =   840
   End
   Begin VB.PictureBox pictScrolllockOn 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2040
      Picture         =   "FrmMain.frx":21D65
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pictScrolllockOff 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1680
      Picture         =   "FrmMain.frx":222EF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pictCapslockOn 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2040
      Picture         =   "FrmMain.frx":22879
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pictCapslockOff 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1680
      Picture         =   "FrmMain.frx":22E03
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pictNumlockOn 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2040
      Picture         =   "FrmMain.frx":2338D
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pictNumlockOff 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1680
      Picture         =   "FrmMain.frx":23917
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pictKeyLock 
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      ToolTipText     =   "NumLock Status"
      Top             =   960
      Width           =   375
   End
   Begin VB.Timer TmrDeKey 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2640
      Top             =   0
   End
   Begin VB.Label LabVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label LabStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Menu mTray 
      Caption         =   "Menu"
      Begin VB.Menu mTrayOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mTraySep1 
         Caption         =   "-"
      End
      Begin VB.Menu mTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private AppInit As Boolean
Private AppExit As Boolean
Private Const RegStartPath As String = "Software\Microsoft\Windows\CurrentVersion\Run"

Private Sub chkHideOnStart_Click()

  If (AppInit = False) Then Exit Sub
  
  SaveConfig

End Sub

Private Sub chkStartWithWin_Click()

  If (chkStartWithWin.Value = 1) Then
    RegWrite HKEY_CURRENT_USER, RegStartPath, "KeyLockStat", Chr(34) & App.Path & "\" & App.EXEName & ".exe" & Chr(34), REG_SZ
  Else
    RegDeleteSetting HKEY_CURRENT_USER, RegStartPath, "KeyLockStat"
  End If

End Sub

Private Sub chkTrayVisible_Click(Index As Integer)

  Dim sTag As String
  Dim bVis As Boolean
  Dim iTrayID As Integer
  Dim sToolTip As String
  
  If (AppInit = False) Then Exit Sub
  
  sTag = chkTrayVisible(Index).Tag
  bVis = True
  
  If (chkTrayVisible(Index).Value = 0) Then bVis = False
  
  iTrayID = GetIconIDFromDesc(sTag)
  
  Select Case sTag
    Case "num": sToolTip = "NumLock Status"
    Case "caps": sToolTip = "CapsLock Status"
    Case "scroll": sToolTip = "ScrollLock Status"
  End Select
  
  If (iTrayID < 0) Then
    If (bVis = True) Then AddIconToTray sTag, sToolTip, pictTrayIcon, pictKeyLock(Index)
  Else
    TrayIconVisible sTag, bVis
  End If
  
  SaveConfig

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  If (KeyCode = vbKeyEscape) Then Unload Me

End Sub

Private Sub Form_Load()

  AppInit = False
  AppExit = False
  
  InitCommonControls
  TrayIconInit
  
  tmrInitApp.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

  If (AppExit = False) Then
    Me.Visible = False
    Cancel = 1
    Exit Sub
  End If
  
  DeleteIconFromTray

End Sub

Private Sub mTrayExit_Click()

  AppExit = True
  Unload Me

End Sub

Private Sub mTrayOptions_Click()

  Me.Visible = True
  Me.WindowState = 0

End Sub

Private Sub pictTrayIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim Message As Long
  
  Message = X / Screen.TwipsPerPixelX
  
  Select Case Message
    Case WM_LBUTTONDBLCLK: Call mTrayOptions_Click
    Case WM_RBUTTONUP: PopupMenu mTray
  End Select

End Sub

Private Sub TmrDeKey_Timer()

  ' Get keystates
  vbKeyState
  
  If (AppInit = False) Then
    ' Remove form editor borders
    LabStatus.BorderStyle = 0
    LabVersion.BorderStyle = 0
    pictKeyLock(0).BorderStyle = 0
    pictKeyLock(1).BorderStyle = 0
    pictKeyLock(2).BorderStyle = 0
  Else
    ' If nothing changed, no need to refresh or update anything
    If (Val(pictKeyLock(0).Tag) = KeyLockState.stNumLock And Val(pictKeyLock(1).Tag) = KeyLockState.stCapsLock And _
      Val(pictKeyLock(2).Tag) <> KeyLockState.stScrollLock) Then Exit Sub
  End If
  
  ' Set tray icon texts
  LabStatus.Caption = "Numlock: " & KeyLockState.szNumLock & Chr(13)
  LabStatus.Caption = LabStatus.Caption & "Capslock: " & KeyLockState.szCapsLock & Chr(13)
  LabStatus.Caption = LabStatus.Caption & "Scrolllock: " & KeyLockState.szScrollLock
  LabVersion.Caption = "Version: " & Format(App.Major) & "." & Format(App.Minor) & "." & Format(App.Revision) & Chr(13)
  LabVersion.Caption = LabVersion.Caption & "Created by: " & App.CompanyName & Chr(13)
  LabVersion.Caption = LabVersion.Caption & "davirichar@gmail.com" & Chr(13)
  LabVersion.Caption = LabVersion.Caption & App.Comments
  
  ' Change tray icon pictures
  If (KeyLockState.stNumLock = 0) Then pictKeyLock(0).Picture = pictNumlockOff.Picture Else pictKeyLock(0).Picture = pictNumlockOn.Picture
  If (KeyLockState.stCapsLock = 0) Then pictKeyLock(1).Picture = pictCapslockOff.Picture Else pictKeyLock(1).Picture = pictCapslockOn.Picture
  If (KeyLockState.stScrollLock = 0) Then pictKeyLock(2).Picture = pictScrolllockOff.Picture Else pictKeyLock(2).Picture = pictScrolllockOn.Picture
  
  ' Add tray icons
  If (AppInit = False) Then
    If (chkTrayVisible(0).Value = 1) Then AddIconToTray "num", "NumLock Status", pictTrayIcon, pictKeyLock(0)
    pictKeyLock(0).Tag = Format(KeyLockState.stNumLock)
    
    If (chkTrayVisible(1).Value = 1) Then AddIconToTray "caps", "CapsLock Status", pictTrayIcon, pictKeyLock(1)
    pictKeyLock(1).Tag = Format(KeyLockState.stCapsLock)
    
    If (chkTrayVisible(2).Value = 1) Then AddIconToTray "scroll", "ScrollLock Status", pictTrayIcon, pictKeyLock(2)
    pictKeyLock(2).Tag = Format(KeyLockState.stScrollLock)
    
    AppInit = True
  End If
  
  ' Update tray icons
  If (Val(pictKeyLock(0).Tag) <> KeyLockState.stNumLock And AppInit = True And chkTrayVisible(0).Value = 1) Then
    UpdateTrayIconPicture "num", pictKeyLock(0)
    pictKeyLock(0).Tag = Format(KeyLockState.stNumLock)
  End If
  
  If (Val(pictKeyLock(1).Tag) <> KeyLockState.stCapsLock And AppInit = True And chkTrayVisible(1).Value = 1) Then
    UpdateTrayIconPicture "caps", pictKeyLock(1)
    pictKeyLock(1).Tag = Format(KeyLockState.stCapsLock)
  End If
  
  If (Val(pictKeyLock(2).Tag) <> KeyLockState.stScrollLock And AppInit = True And chkTrayVisible(2).Value = 1) Then
    UpdateTrayIconPicture "scroll", pictKeyLock(2)
    pictKeyLock(2).Tag = Format(KeyLockState.stScrollLock)
  End If

End Sub

Private Sub ReadConfig()

  Dim sConfigFile As String
  Dim sValue As String
  
  sConfigFile = GetFilePath("KeyLockStat_config.ini")
  
  If (FileExists(sConfigFile) = False) Then Exit Sub
  
  sValue = ReadIniSetting("VisibleIcons", "num", sConfigFile)
  If (Val(sValue) = 0) Then chkTrayVisible(0).Value = 0
  
  sValue = ReadIniSetting("VisibleIcons", "caps", sConfigFile)
  If (Val(sValue) = 0) Then chkTrayVisible(1).Value = 0
  
  sValue = ReadIniSetting("VisibleIcons", "scroll", sConfigFile)
  If (Val(sValue) = 0) Then chkTrayVisible(2).Value = 0
  
  sValue = ReadIniSetting("General", "HideOnStart", sConfigFile)
  If (Val(sValue) = 1) Then chkHideOnStart.Value = 1
  
  sValue = RegRead(HKEY_CURRENT_USER, RegStartPath, "KeyLockStat")
  If (sValue = Chr(34) & App.Path & "\" & App.EXEName & ".exe" & Chr(34)) Then chkStartWithWin.Value = 1

End Sub

Private Sub SaveConfig()
  
  Dim sConfigFile As String
  Dim sValue As String
  
  sConfigFile = GetFilePath("KeyLockStat_config.ini")
  
  sValue = Format(chkTrayVisible(0).Value)
  SaveIniSetting "VisibleIcons", "num", sValue, sConfigFile
  
  sValue = Format(chkTrayVisible(1).Value)
  SaveIniSetting "VisibleIcons", "caps", sValue, sConfigFile
  
  sValue = Format(chkTrayVisible(2).Value)
  SaveIniSetting "VisibleIcons", "scroll", sValue, sConfigFile
  
  sValue = Format(chkHideOnStart.Value)
  SaveIniSetting "General", "HideOnStart", sValue, sConfigFile
  
End Sub

Private Sub tmrInitApp_Timer()

  tmrInitApp.Enabled = False
  
  ReadConfig
  
  If (chkHideOnStart.Value = 0) Then Me.Visible = True
  
  TmrDeKey.Enabled = True

End Sub
