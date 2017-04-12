VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommPortList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private listCount As Long
Private arr() As SerialDevice

Private wmi As WbemScripting.SWbemServices

Private Type SerialDevice
    friendlyName As String
    commPort As String
    classGUID As String
    instanceName As String
    locationInformation As String
End Type


Private Sub Class_Initialize()
    Clear
    
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\WMI") '\\.\root\cimv2
End Sub

Public Sub Clear()
    ReDim arr(0)
    listCount = 0
End Sub

Public Property Get Count() As Long
    Count = listCount
End Property

Public Property Get friendlyName(index As Long) As String
    If listCount = 0 Then Exit Property
    friendlyName = arr(index).friendlyName
End Property

Public Property Get commPort(index As Long) As String
    If listCount = 0 Then Exit Property
    commPort = arr(index).commPort
End Property

Public Property Get locationInformation(index As Long) As String
    If listCount = 0 Then Exit Property
    locationInformation = arr(index).locationInformation
End Property



Public Sub Refresh()
    Dim colItems As Object, objItem As Object
    Dim itemCount As Long
    Dim i As Long
    Dim singleDevice As SerialDevice
    On Error Resume Next
    
    Clear
    Set colItems = wmi.ExecQuery("SELECT * from MSSerial_PortName", , 48) ' WHERE Name LIKE '%(COM%' ''' 'Win32_SerialPort
    
    For Each objItem In colItems
        If objItem Is Nothing Then Exit Sub
        
        ReDim Preserve arr(0 To listCount)
        
        With arr(listCount)
            .commPort = objItem.PortName
            .instanceName = objItem.instanceName
        End With
        
        fillDeviceProperties listCount
        
        listCount = listCount + 1
    Next
End Sub

Private Sub fillDeviceProperties(index As Long) 'SerialDevice
    Dim hKey As Long
    Dim Length As Long
    Dim Value As String
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Enum\" & Left(arr(index).instanceName, Len(arr(index).instanceName) - 2), 0&, KEY_READ, hKey) <> ERROR_SUCCESS Then
        Exit Sub
    End If
    
    '##############################
    Length = 0
    RegQueryValueEx hKey, "FriendlyName", 0&, REG_SZ, 0&, Length
    Value = Space(Length)
    RegQueryValueEx hKey, "FriendlyName", 0&, REG_SZ, ByVal Value, Length
    
    arr(index).friendlyName = Trim(Replace(Value, Chr(0), ""))
    '##############################
    
    
    '##############################
    Length = 0
    RegQueryValueEx hKey, "LocationInformation", 0&, REG_SZ, 0&, Length
    Value = Space(Length)
    RegQueryValueEx hKey, "LocationInformation", 0&, REG_SZ, ByVal Value, Length
    
    arr(index).locationInformation = Trim(Replace(Value, Chr(0), ""))
    '##############################
    
    
    '##############################
    Length = 0
    RegQueryValueEx hKey, "ClassGUID", 0&, REG_SZ, 0&, Length
    Value = Space(Length)
    RegQueryValueEx hKey, "ClassGUID", 0&, REG_SZ, ByVal Value, Length
    
    arr(index).classGUID = Trim(Replace(Value, Chr(0), ""))
    '##############################
    
    
    RegCloseKey (hKey)
End Sub
