Attribute VB_Name = "Registry"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const REG_SZ As Long = 1
Global Const ERROR_MORE_DATA = 234
Global Const ERROR_SUCCESS As Long = 0
Global Const LB_SETTABSTOPS As Long = &H192

Global Const STANDARD_RIGHTS_READ As Long = &H20000
Global Const KEY_QUERY_VALUE As Long = &H1
Global Const KEY_ALL_ACCESS = &H3F
Global Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Global Const KEY_NOTIFY As Long = &H10
Global Const SYNCHRONIZE As Long = &H100000
Global Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or _
                                   KEY_QUERY_VALUE Or _
                                   KEY_ENUMERATE_SUB_KEYS Or _
                                   KEY_NOTIFY) And _
                                   (Not SYNCHRONIZE))

Const LVM_FIRST As Long = &H1000
Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Type REGISTRY_APPINFO
   RegistryName As String
   DisplayName As String
   DisplayVersion As String
   CanUninstall As Boolean
   UninstallString As String
End Type

Type FILETIME 'ft
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Declare Function RegOpenKeyEx Lib "advapi32.dll" _
   Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, _
   ByVal lpSubKey As String, _
   ByVal ulOptions As Long, _
   ByVal samDesired As Long, _
   phkResult As Long) As Long
   
Declare Function RegEnumKeyEx Lib "advapi32.dll" _
   Alias "RegEnumKeyExA" _
   (ByVal hKey As Long, _
   ByVal dwIndex As Long, _
   ByVal lpName As String, _
   lpcbName As Long, _
   ByVal lpReserved As Long, _
   ByVal lpClass As String, _
   lpcbClass As Long, _
   lpftLastWriteTime As FILETIME) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" _
   Alias "RegQueryValueExA" _
  (ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   lpType As Long, _
   lpData As Any, _
   lpcbData As Long) As Long
   
Declare Function RegQueryInfoKey Lib "advapi32.dll" _
   Alias "RegQueryInfoKeyA" _
  (ByVal hKey As Long, _
   ByVal lpClass As String, _
   lpcbClass As Long, _
   ByVal lpReserved As Long, _
   lpcSubKeys As Long, _
   lpcbMaxSubKeyLen As Long, _
   lpcbMaxClassLen As Long, _
   lpcValues As Long, _
   lpcbMaxValueNameLen As Long, _
   lpcbMaxValueLen As Long, _
   lpcbSecurityDescriptor As Long, _
   lpftLastWriteTime As FILETIME) As Long

Declare Function RegCloseKey Lib "advapi32.dll" _
  (ByVal hKey As Long) As Long

Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long

Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long


Function getCommPortList(ByRef resultList() As String)
    Dim subkeys As Collection
    Dim subkey_values As Collection
    Dim subkey_num As Integer
    Dim subkey_name As String
    Dim subkey_value As String
    Dim Length As Long
    Dim hKey As Long
    Dim txt As String
    Dim subkey_txt As String
    Dim value_num As Long
    Dim value_name_len As Long
    Dim value_name As String
    Dim reserved As Long
    Dim value_type As Long
    Dim value_string As String
    Dim value_data(1 To 1024) As Byte
    Dim value_data_len As Long
    Dim i As Integer
    
    ReDim resultList(0) As String
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\SERIALCOMM", 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
        getCommPortList = ""
        Exit Function
    End If
    
    value_num = 0
    
    Do
        
        value_name_len = 1024
        value_name = Space$(value_name_len)
        value_data_len = 1024

        If RegEnumValue(hKey, value_num, value_name, value_name_len, 0, value_type, value_data(1), value_data_len) <> ERROR_SUCCESS Then Exit Do
        
        value_name = Left$(value_name, value_name_len)
        
        
        'Debug.Print value_name
        Select Case value_type
            Case REG_SZ
                    value_string = ""
                    For i = 1 To value_data_len - 1
                        value_string = value_string & _
                            Chr$(value_data(i))
                    Next i
                    
                    resultList(value_num) = value_string
        End Select
        
        value_num = value_num + 1
    Loop
    
    'Debug.Print txt
    
End Function



