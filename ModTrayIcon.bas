Attribute VB_Name = "ModTrayIcon"
Option Explicit

Private Const NOTIFYICON_VERSION = &H3

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10

Private Const NOTIFYICONDATA_V1_SIZE As Long = 88
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504
Private NOTIFYICONDATA_SIZE As Long
   
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private atiData() As NOTIFYICONDATA
Private atiIds() As String

Private Const colBG = vbMagenta

' Initialises the Tray Icon array
Public Sub TrayIconInit()
  
  ReDim atiData(0)
  ReDim atiIds(0)
  
End Sub

' Update Tray icon picture
Public Sub UpdateTrayIconPicture(ByVal sIconDesc As String, ByVal oPictBox As PictureBox)
  
  Dim icoID As Integer
  
  icoID = GetIconIDFromDesc(sIconDesc)
  
  If (icoID < 0) Then
    MsgBox "Cannot update icon picture.", vbOKOnly + vbCritical, "Negative icon index"
    Exit Sub
  End If
  
  atiData(icoID).hIcon = oPictBox.Picture
  
  Call Shell_NotifyIcon(NIM_MODIFY, atiData(icoID))
  
End Sub

' Showor Hide tray icon
Public Sub TrayIconVisible(ByVal sIconDesc As String, ByVal bVisible As Boolean)
  
  Dim lVisible As Long
  Dim icoID As Integer
  Dim atIcon As NOTIFYICONDATA
  
  icoID = GetIconIDFromDesc(sIconDesc)
  
  atIcon = atiData(icoID)
  
  If (bVisible = True) Then
    Call Shell_NotifyIcon(NIM_ADD, atIcon)
  Else
    atIcon.uCallbackMessage = 0
    Call Shell_NotifyIcon(NIM_DELETE, atIcon)
  End If
  
End Sub

' Adds an icon to the system tray
Public Sub AddIconToTray(ByVal sIconDesc As String, ByVal sToolTipText As String, ByVal oPictBox As PictureBox, ByVal oPictBoxIcon As PictureBox)
  
  Dim itiID As Integer
  
  itiID = 0
  
  If (atiData(itiID).uID > 0) Then
    ReDim Preserve atiData(UBound(atiData) + 1)
    ReDim Preserve atiIds(UBound(atiData))
    itiID = UBound(atiData)
  End If
  
  If (sIconDesc = "") Then
    MsgBox "Icon description cannot be empty. Tray icon has NOT been added.", vbOKOnly + vbCritical, "No description"
    Exit Sub
  End If
  
  ' Add an icon description for ease of search and access
  atiIds(itiID) = sIconDesc
  
  atiData(itiID).cbSize = Len(atiData(itiID))
  atiData(itiID).hwnd = oPictBox.hwnd
  atiData(itiID).uID = itiID + 1
  atiData(itiID).uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  atiData(itiID).uCallbackMessage = WM_MOUSEMOVE
  atiData(itiID).hIcon = oPictBoxIcon.Picture
  atiData(itiID).szTip = sToolTipText & vbNullChar
  
  Call Shell_NotifyIcon(NIM_ADD, atiData(itiID))

End Sub

' Delete and icon from tray area
Public Sub DeleteIconFromTray(Optional ByVal sIconDesc As String = "")

  Dim icoID As Integer
  Dim iCtr As Integer
  
  icoID = GetIconIDFromDesc(sIconDesc)
  
  ' If no Icon ID and description is given remove all tray icons
  If (icoID < 0 And sIconDesc = "") Then
    For iCtr = 0 To UBound(atiData)
      Call Shell_NotifyIcon(NIM_DELETE, atiData(iCtr))
    Next
  End If
  
  ' Delete tray icon if an ID is found
  If (icoID >= 0) Then Call Shell_NotifyIcon(NIM_DELETE, atiData(icoID))
  
End Sub

' Get the icon ID from it's description
Public Function GetIconIDFromDesc(ByVal sIconDesc As String) As Integer
  
  Dim iCtr As Integer
  Dim RetVal As Integer
  
  RetVal = -1
  
  For iCtr = 0 To UBound(atiIds)
    If (atiIds(iCtr) = sIconDesc) Then
      RetVal = iCtr
      Exit For
    End If
  Next
  
  GetIconIDFromDesc = RetVal
  
End Function
