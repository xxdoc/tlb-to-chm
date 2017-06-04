VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' 网上找到的一个注册表的操作类

'注册表主键
Public Enum enumRegMainKey
   iHKEY_CURRENT_USER = &H80000001
   iHKEY_LOCAL_MACHINE = &H80000002
   iHKEY_CLASSES_ROOT = &H80000000
   iHKEY_CURRENT_CONFIG = &H80000005
   iHKEY_USERS = &H80000003
End Enum
'注册表数据类型
Public Enum enumRegSzType
   iREG_SZ = &H1
   iREG_EXPAND_SZ = &H2
   iREG_BINARY = &H3
   iREG_DWORD = &H4
   iREG_NONE = 0&
   iREG_DWORD_LITTLE_ENDIAN = 4&
   iREG_DWORD_BIG_ENDIAN = 5&
   iREG_LINK = 6&
   iREG_MULTI_SZ = 7&
   iREG_RESOURCE_LIST = 8&
   iREG_FULL_RESOURCE_DEscriptOR = 9&
   iREG_RESOURCE_REQUIREMENTS_LIST = 10&
End Enum
'注册表
Private Const ERROR_SUCCESS As Long = 0&
'Private Const ERROR_BADDB = 1009&
'Private Const ERROR_BADKEY = 1010&
'Private Const ERROR_CANTOPEN = 1011&
'Private Const ERROR_CANTREAD = 1012&
'Private Const ERROR_CANTWRITE = 1013&
'Private Const ERROR_OUTOFMEMORY = 14&
'Private Const ERROR_INVALID_PARAMETER = 87&
'Private Const ERROR_ACCESS_DENIED = 5&
'Private Const ERROR_NO_MORE_ITEMS = 259&
'Private Const ERROR_MORE_DATA = 234&
Private Const KEY_QUERY_VALUE As Long = &H1&
Private Const KEY_SET_VALUE As Long = &H2&
Private Const KEY_CREATE_SUB_KEY As Long = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8&
Private Const KEY_NOTIFY As Long = &H10&
'Private Const KEY_CREATE_LINK = &H20&
'Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL As Long = &H20000
'Private Const WRITE_DAC = &H40000
'Private Const WRITE_OWNER = &H80000
'Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
'Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const KEY_READ As Long = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE As Long = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
'Private Const KEY_EXECUTE = KEY_READ
'Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
'Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long       ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal lpSecurityAttributes As Long) As Long
'Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwflags As Long) As Long
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Boolean
End Type
Public Function GetValue(ByVal mainKey As enumRegMainKey, _
                  ByVal subKey As String, _
                  ByVal keyV As String, _
                  ByRef sValue As Variant, _
                  Optional ByRef rlngErrNum As Long, _
                  Optional ByRef rstrErrDescr As String) As Boolean
   Dim hKey As Long, lType As Long, lBuffer As Long, sBuffer As String, lData As Long
   On Error GoTo GetValueErr
   GetValue = False
   If RegOpenKeyEx(mainKey, subKey, 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while fetching registry key."
   End If
   If RegQueryValueEx(hKey, keyV, 0, lType, ByVal 0, lBuffer) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while fetching registry key."
   End If
   Select Case lType
     Case iREG_SZ
       lBuffer = 255
       sBuffer = Space(lBuffer)
       If RegQueryValueEx(hKey, keyV, 0, lType, ByVal sBuffer, lBuffer) <> ERROR_SUCCESS Then
         Err.Raise vbObjectError + 1, , "Error accured while fetching registry key."
       End If
       sValue = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
     Case iREG_EXPAND_SZ
       sBuffer = Space(lBuffer)
       If RegQueryValueEx(hKey, keyV, 0, lType, ByVal sBuffer, lBuffer) <> ERROR_SUCCESS Then
         Err.Raise vbObjectError + 1, , "Error accured while fetching registry key."
       End If
       sValue = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
     Case iREG_DWORD
       If RegQueryValueEx(hKey, keyV, 0, lType, lData, lBuffer) <> ERROR_SUCCESS Then
         Err.Raise vbObjectError + 1, , "Error accured while fetching registry key."
       End If
       sValue = lData
     Case iREG_BINARY
       If RegQueryValueEx(hKey, keyV, 0, lType, lData, lBuffer) <> ERROR_SUCCESS Then
         Err.Raise vbObjectError + 1, , "Error accured while fetching registry key."
       End If
       sValue = lData
   End Select
   If RegCloseKey(hKey) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while fetching registry key."
   End If
   GetValue = True
   Err.Clear
GetValueErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
End Function
Public Function SetValue(ByVal mainKey As enumRegMainKey, _
                  ByVal subKey As String, _
                  ByVal keyV As String, _
                  ByVal lType As enumRegSzType, _
                  ByVal sValue As Variant, _
                  Optional ByRef rlngErrNum As Long, _
                  Optional ByRef rstrErrDescr As String) As Boolean
   Dim S As Long, lBuffer As Long, hKey As Long
   Dim ss As SECURITY_ATTRIBUTES
   On Error GoTo SetValueErr
   SetValue = False
   ss.nLength = Len(ss)
   ss.lpSecurityDescriptor = 0
   ss.bInheritHandle = True
   If RegCreateKeyEx(mainKey, subKey, 0, "", 0, KEY_WRITE, ss, hKey, S) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while writing registry key."
   End If
   Select Case lType
     Case iREG_SZ
       lBuffer = LenB(sValue)
       If RegSetValueEx(hKey, keyV, 0, lType, ByVal sValue, lBuffer) <> ERROR_SUCCESS Then
         Err.Raise vbObjectError + 1, , "Error accured while writing registry key."
       End If
     Case iREG_EXPAND_SZ
       lBuffer = LenB(sValue)
       If RegSetValueEx(hKey, keyV, 0, lType, ByVal sValue, lBuffer) <> ERROR_SUCCESS Then
         Err.Raise vbObjectError + 1, , "Error accured while writing registry key."
       End If
     Case iREG_DWORD
       lBuffer = 4
       If RegSetValueExA(hKey, keyV, 0, lType, sValue, lBuffer) <> ERROR_SUCCESS Then
         Err.Raise vbObjectError + 1, , "Error accured while writing registry key."
       End If
     Case iREG_BINARY
       lBuffer = 4
       If RegSetValueExA(hKey, keyV, 0, lType, sValue, lBuffer) <> ERROR_SUCCESS Then
         Err.Raise vbObjectError + 1, , "Error accured while writing registry key."
       End If
     Case Else
       Err.Raise vbObjectError + 1, , "Unsupported parameter(s)"
   End Select
   If RegCloseKey(hKey) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while writing registry key."
   End If
   SetValue = True
   Err.Clear
SetValueErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
End Function
Public Function DeleteValue(ByVal mainKey As enumRegMainKey, _
                     ByVal subKey As String, _
                     ByVal keyV As String, _
                     Optional ByRef rlngErrNum As Long, _
                     Optional ByRef rstrErrDescr As String) As Boolean
   Dim hKey As Long
   On Error GoTo DeleteValueErr
   DeleteValue = False
   If RegOpenKeyEx(mainKey, subKey, 0, KEY_WRITE, hKey) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while deleting registry value"
   End If
   If RegDeleteValue(hKey, keyV) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while deleting registry value"
   End If
   If RegCloseKey(hKey) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while deleting registry value"
   End If
   DeleteValue = True
   Err.Clear
DeleteValueErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
End Function
Public Function DeleteKey(ByVal mainKey As enumRegMainKey, _
                     ByVal subKey As String, _
                     ByVal keyV As String, _
                     Optional ByRef rlngErrNum As Long, _
                     Optional ByRef rstrErrDescr As String) As Boolean
   Dim hKey As Long
   On Error GoTo DeleteKeyErr
   DeleteKey = False
   If RegOpenKeyEx(mainKey, subKey, 0, KEY_WRITE, hKey) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while deleting registry value"
   End If
   If RegDeleteKey(hKey, keyV) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while deleting registry value"
   End If
   If RegCloseKey(hKey) <> ERROR_SUCCESS Then
     Err.Raise vbObjectError + 1, , "Error accured while deleting registry value"
   End If
   DeleteKey = True
   Err.Clear
DeleteKeyErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
End Function

Public Function GetSubKeys(ByVal mainKey As enumRegMainKey, ByVal subKey As String) As Collection
   Set GetSubKeys = New Collection

   Dim hKey As Long
   If RegOpenKeyEx(mainKey, subKey, 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
      Exit Function
   End If
   
   Dim i As Long
   i = 0
   
   Do
      Const buffer_size As Long = 100
   
      Dim buffer As String
      buffer = Space(buffer_size)
      
      Dim ft As FILETIME
      If RegEnumKeyEx(hKey, i, buffer, buffer_size, 0, vbNullString, 0, ft) <> ERROR_SUCCESS Then
         Exit Do
      End If
      
      GetSubKeys.Add Trim(Replace(buffer, Chr(0), ""))
      
      i = i + 1
   Loop While True
   
   RegCloseKey hKey
End Function

Private Sub Class_Initialize()
   LogInitialize Me
End Sub

Private Sub Class_Terminate()
   LogTerminate Me
End Sub

