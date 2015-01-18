Attribute VB_Name = "modConfig"
Option Explicit
Public INIFILE As String
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const MAX_PATH = 256&
Public Const REG_SZ = 1
Public Declare Function RegCreateKey& Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, lphKey As Long)
Public Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpszSubKey As String, ByVal fdwType As Long, ByVal lpszValue As String, ByVal dwLength As Long)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function INIRead(Session As String, KeyValue As String, INIFILE As String) As String
Dim s As String * 1024
Dim ReturnValue As Long
ReturnValue = GetPrivateProfileString(Session, KeyValue, "", s, 1024, INIFILE)
INIRead = Left$(s, InStr(s, Chr(0)) - 1)
End Function
 
Public Function INIWrite(Session As String, KeyValue As String, DataValue As String, INIFILE As String) As String
Dim ReturnValue As Long
ReturnValue = WritePrivateProfileString(Session, KeyValue, DataValue, INIFILE)
End Function

Public Function SetDefExt(AppName As String, Description As String, Extension As String, AppPath As String)
On Error Resume Next
Dim ret As Long
Dim lphKey As Long
Dim FilePath As String

ret = RegCreateKey&(HKEY_CLASSES_ROOT, AppName, lphKey)
ret = RegSetValue&(lphKey&, Empty, REG_SZ, Description, 0&)
ret = RegCreateKey&(HKEY_CLASSES_ROOT, Extension, lphKey)
ret = RegSetValue&(lphKey, Empty, REG_SZ, AppName, 0&)
ret = RegCreateKey&(HKEY_CLASSES_ROOT, AppName, lphKey)
ret = RegSetValue&(lphKey, "shell\open\command", REG_SZ, AppPath & " %1", MAX_PATH)

If Not ret = 0 Then
    MsgBox "Need Administrator Permission", vbCritical, "Permission Error"
    End
    Exit Function
End If
End Function
