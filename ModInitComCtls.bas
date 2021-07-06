Attribute VB_Name = "ModInitComCtls"
Option Explicit

Private Type tpInitCommonControlsEx
  lngSize As Long
  lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tpInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControls() As Boolean
  
  On Error Resume Next
  Dim iccex As tpInitCommonControlsEx
  
  With iccex
    .lngSize = LenB(iccex)
    .lngICC = ICC_USEREX_CLASSES
  End With
  InitCommonControlsEx iccex
  InitCommonControls = (Err.Number = 0)
  On Error GoTo 0
  
End Function

