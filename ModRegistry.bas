Attribute VB_Name = "ModRegistry"
Option Explicit

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Public Const HKEY_PERFORMANCE_TEXT As Long = &H80000050
Public Const HKEY_PERFORMANCE_NLSTEXT As Long = &H80000060
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_DYN_DATA As Long = &H80000006
Public Const HKEY_CURRENT_USER_LOCAL_SETTINGS As Long = &H80000007

Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4

Private Const KEY_QUERY_VALUE As Long = 1
Private Const KEY_SET_VALUE As Long = 2
Private Const KEY_CREATE_SUB_KEY As Long = 4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = 8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20
Private Const KEY_WOW64_64KEY As Long = &H100
Private Const KEY_WOW64_32KEY As Long = &H200
Private Const KEY_ALL_ACCESS As Long = &HF003F
Private Const KEY_WRITE As Long = &H20006
Private Const KEY_EXECUTE As Long = &H20019
Private Const KEY_READ As Long = &H20019

Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_MORE_DATA As Long = 234

' Read value from registry
Public Function RegRead(ByVal lhKey As Long, ByVal sPath As String, ByVal sSetting As String) As String

  Dim hKey As Long
  Dim lRetVal As Long
  Dim lDataType As Long
  Dim lLen As Long
  Dim RetVal As String
  
  lRetVal = RegOpenKeyEx(lhKey, sPath, 0, KEY_QUERY_VALUE, hKey)
  lRetVal = RegQueryValueEx(hKey, sSetting, 0, lDataType, ByVal RetVal, lLen)
  
  If (lRetVal = ERROR_MORE_DATA Or Len(RetVal) < lLen - 1) Then
    RetVal = Space(lLen)
    lRetVal = RegQueryValueEx(hKey, sSetting, 0, 0, ByVal RetVal, lLen)
  End If
  
  If (lDataType = REG_SZ) Then RetVal = Left(RetVal, lLen - 1)
  
  lRetVal = RegCloseKey(hKey)
  
  RegRead = RetVal

End Function

' Write value in the registry
Public Sub RegWrite(ByVal lhKey As Long, ByVal sPath As String, ByVal sSetting As String, ByVal sValue As String, ByVal lValType As Long)

  Dim hKey As Long
  Dim lRetVal As Long
  Dim lLen As Long
  Dim lNewOpened As Long
  Dim SecAttr As SECURITY_ATTRIBUTES
  
  SecAttr.nLength = Len(SecAttr)
  SecAttr.lpSecurityDescriptor = 0
  SecAttr.bInheritHandle = 1
  
  lRetVal = RegCreateKeyEx(lhKey, sPath, 0, "", 0, KEY_WRITE, SecAttr, hKey, lNewOpened)
  If (lRetVal <> 0) Then Exit Sub
  
  lRetVal = RegSetValueEx(hKey, sSetting, 0, lValType, ByVal sValue & vbNullChar, Len(sValue))
  lRetVal = RegCloseKey(hKey)

End Sub

' Delete value from the registry
Public Sub RegDeleteSetting(ByVal lhKey As Long, ByVal sPath As String, ByVal sSetting As String)

  Dim lRetVal As Long
  Dim hKey As Long
  
  lRetVal = RegOpenKeyEx(lhKey, sPath, 0, KEY_QUERY_VALUE, hKey)
  lRetVal = RegDeleteValue(hKey, sSetting)
  lRetVal = RegCloseKey(hKey)

End Sub
