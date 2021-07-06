Attribute VB_Name = "ModINI"
Option Explicit

Private Const lParamLength As Integer = 2

Public Declare Function GetPrivateProfileString Lib "kernel32" _
  Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
  Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long

' Get a file path from a filename or from a full path
Public Function GetFilePath(Optional ByVal sFileName As String = "", Optional ByVal sFileFullPath As String = "") As String

  Dim RetVal As String
  
  If (sFileName = "") Then
    RetVal = sFileFullPath
  Else
    RetVal = App.Path & "\" & sFileName
  End If
  
  GetFilePath = RetVal

End Function

' Read setting from a INI file
Public Function ReadIniSetting(ByVal sHeading As String, ByVal sKey As String, ByVal sFilePath As String) As String
  
  Dim sRetVal As String * lParamLength
  Dim sDefault As String * lParamLength
  Dim lLen As Long
  
  lLen = GetPrivateProfileString(sHeading, sKey, sDefault, sRetVal, lParamLength, sFilePath)
  
  ReadIniSetting = Mid(sRetVal, 1, lLen)
  
End Function

' Write a setting to a INI file
Public Sub SaveIniSetting(ByVal sHeading As String, ByVal sKey As String, ByVal sValue As String, ByVal sFilePath As String)

  Dim sRetVal As String * lParamLength
  Dim sDefault As String * lParamLength
  Dim lLen As Long
  
  lLen = WritePrivateProfileString(sHeading, sKey, sValue, sFilePath)

End Sub
