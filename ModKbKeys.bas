Attribute VB_Name = "ModKbKeys"
Option Explicit

Public Type tpKeyStates
  stNumLock As Integer
  stCapsLock As Integer
  stScrollLock As Integer
  stInsert As Integer
  szNumLock As String
  szCapsLock As String
  szScrollLock As String
  szInsert As String
End Type

Public KeyLockState As tpKeyStates

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Function vbKeyState(Optional ByVal sKeyType As String = "") As Integer
  
  Dim sKType As String
  Dim RetVal As Integer
  Dim iKStat As Integer
  
  sKType = LCase(sKeyType)
  iKStat = -1
  
  Select Case sKType
    Case "num":
      KeyLockState.stNumLock = GetKeyState(vbKeyNumlock)
      iKStat = KeyLockState.stNumLock
    Case "caps":
      KeyLockState.stCapsLock = GetKeyState(vbKeyCapital)
      iKStat = KeyLockState.stCapsLock
    Case "scroll":
      KeyLockState.stScrollLock = GetKeyState(vbKeyScrollLock)
      iKStat = KeyLockState.stScrollLock
    Case "insert":
      KeyLockState.stInsert = GetKeyState(vbKeyInsert)
      iKStat = KeyLockState.stInsert
    Case "":
      KeyLockState.stNumLock = GetKeyState(vbKeyNumlock)
      KeyLockState.stCapsLock = GetKeyState(vbKeyCapital)
      KeyLockState.stScrollLock = GetKeyState(vbKeyScrollLock)
      KeyLockState.stInsert = GetKeyState(vbKeyInsert)
  End Select
  
  If (KeyLockState.stNumLock = 1) Then KeyLockState.szNumLock = "ON" Else KeyLockState.szNumLock = "OFF"
  If (KeyLockState.stCapsLock = 1) Then KeyLockState.szCapsLock = "ON" Else KeyLockState.szCapsLock = "OFF"
  If (KeyLockState.stScrollLock = 1) Then KeyLockState.szScrollLock = "ON" Else KeyLockState.szScrollLock = "OFF"
  If (KeyLockState.stInsert = 1) Then KeyLockState.szInsert = "ON" Else KeyLockState.szInsert = "OFF"
  
  RetVal = iKStat
  
  vbKeyState = RetVal
  
End Function

