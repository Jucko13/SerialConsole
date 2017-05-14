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

Private ListCount As Long
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
    
    Set wmi = GetObject("winmgmts:\\.\root\cimv2") '"winmgmts:{impersonationLevel=impersonate}!\\.\root\WMI"           '\\.\root\cimv2
End Sub

Public Sub Clear()
    ReDim arr(0)
    ListCount = 0
End Sub

Public Property Get Count() As Long
    Count = ListCount
End Property

Public Property Get friendlyName(Index As Long) As String
    If ListCount = 0 Then Exit Property
    friendlyName = arr(Index).friendlyName
End Property

Public Property Get commPort(Index As Long) As String
    If ListCount = 0 Or Index < 0 Or Index >= ListCount Then Exit Property
    commPort = arr(Index).commPort
End Property

Public Property Get locationInformation(Index As Long) As String
    If ListCount = 0 Then Exit Property
    locationInformation = arr(Index).locationInformation
End Property




'
'Public Sub Refresh()
'    Dim colItems As Object, objItem As Object
'    Dim itemCount As Long
'    Dim i As Long
'    Dim singleDevice As SerialDevice
'    On Error Resume Next
'
'    Clear
'    Set colItems = wmi.ExecQuery("SELECT * from MSSerial_PortName", , 48) ' WHERE Name LIKE '%(COM%' ''' 'Win32_SerialPort
'
'    For Each objItem In colItems
'        If objItem Is Nothing Then Exit Sub
'
'        ReDim Preserve arr(0 To listCount)
'
'        With arr(listCount)
'            .commPort = objItem.PortName
'            .instanceName = objItem.instanceName
'        End With
'
'        fillDeviceProperties listCount
'
'        listCount = listCount + 1
'    Next
'End Sub

Public Property Get isCommAvailable(Index As Long) As Boolean
    If ListCount = 0 Then Exit Property

    Dim k As PerformanceTimer
    Dim totalDevices As Long
    Dim tmpStr As String
    Dim colItems As Object
    
    Set k = New PerformanceTimer
    
    k.StartTimer
    
    'Dim colItem As Object
    
    Set colItems = wmi.ExecQuery("SELECT Availability from Win32_PnPEntity WHERE DeviceID='" & Replace(arr(Index).instanceName, "\", "\\") & "'")
    
    'GetMembers frmMain.List1, colItems
    
    isCommAvailable = colItems.Count > 0
    
    'k.StopTimer
    'Debug.Print k.TimeElapsed(pvMilliSecond)
    
    'Debug.Print tmpStr
End Property


Public Sub Refresh()
    Dim colItems As Object, objItem As Object
    Dim itemCount As Long
    Dim i As Long
    Dim singleDevice As SerialDevice
    On Error Resume Next
    

    
    'k.StartTimer
    Clear
    Set colItems = wmi.ExecQuery("SELECT * from Win32_PnPEntity", , 48) ', , 48 = return immediately ' WHERE Name LIKE '%(COM%' ''' 'Win32_SerialPort
    
    'k.StopTimer
    
    'Debug.Print k.TimeElapsed(pvMilliSecond)
    
    For Each objItem In colItems
        If objItem Is Nothing Then Exit Sub
        
        ReDim Preserve arr(0 To ListCount)
         
        With arr(ListCount)
            .commPort = ""
            .instanceName = ""
        
            '.commPort = objItem.DeviceID
            .instanceName = objItem.deviceid
        End With
        
        
        If fillDeviceProperties(ListCount) Then
            ListCount = ListCount + 1
        End If
    Next
    
    
End Sub

Private Function fillDeviceProperties(Index As Long) As Boolean 'SerialDevice
    Dim hKey As Long, hKey2 As Long
    Dim Length As Long
    Dim Value As String
    
    fillDeviceProperties = False
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Enum\" & arr(Index).instanceName, 0&, KEY_READ, hKey) <> ERROR_SUCCESS Then
        GoTo CloseItAll
    End If
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Enum\" & arr(Index).instanceName & "\Device Parameters", 0&, KEY_READ, hKey2) <> ERROR_SUCCESS Then
        GoTo CloseItAll
    End If
    
    'Debug.Assert Left(arr(index).instanceName, 3) <> "USB"
     
    '##############################
    Length = 0
    RegQueryValueEx hKey2, "PortName", 0&, REG_SZ, 0&, Length
    Value = Space(Length)
    RegQueryValueEx hKey2, "PortName", 0&, REG_SZ, ByVal Value, Length
    
    arr(Index).commPort = Trim(Replace(Value, Chr(0), ""))
    If arr(Index).commPort = "" Then
        GoTo CloseItAll
    End If
    
    '##############################
    
    
    '##############################
    Length = 0
    RegQueryValueEx hKey, "FriendlyName", 0&, REG_SZ, 0&, Length
    Value = Space(Length)
    RegQueryValueEx hKey, "FriendlyName", 0&, REG_SZ, ByVal Value, Length
    
    arr(Index).friendlyName = Trim(Replace(Value, Chr(0), ""))
    '##############################
    
    
    '##############################
    Length = 0
    RegQueryValueEx hKey, "LocationInformation", 0&, REG_SZ, 0&, Length
    Value = Space(Length)
    RegQueryValueEx hKey, "LocationInformation", 0&, REG_SZ, ByVal Value, Length
    
    arr(Index).locationInformation = Trim(Replace(Value, Chr(0), ""))
    '##############################
    
    
    '##############################
    Length = 0
    RegQueryValueEx hKey, "ClassGUID", 0&, REG_SZ, 0&, Length
    Value = Space(Length)
    RegQueryValueEx hKey, "ClassGUID", 0&, REG_SZ, ByVal Value, Length
    
    arr(Index).classGUID = Trim(Replace(Value, Chr(0), ""))
    '##############################
    
    fillDeviceProperties = True
    
CloseItAll:
    
    RegCloseKey (hKey)
End Function
