Attribute VB_Name = "MultipleInstances"
Option Explicit


Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const SW_SHOWNORMAL = 1
Private Const WM_CLOSE = &H10
Private Const WM_SETTEXT = &HC

Private searchSpecificCom As Boolean
Private searchWindowNamePart As String
Private EnumWindowsHWND As Long

Sub SendMessageToInstances(whatToSend As String)
    Dim WinWnd As Long, Ret As String, RetVal As Long, lpClassName As String
    Dim textboxHandle As Long
    
    Dim strSplit() As String
    strSplit = Split(whatToSend, " ")
    
    If UBound(strSplit) <> 1 Then
        MsgBox "Possible commands:" & vbCrLf & vbCrLf & "CC COM4" & vbCrLf & "OC COM5" & vbCrLf & "CC {serial.port}"
        Exit Sub
    End If
    
    searchSpecificCom = (InStr(1, LCase(strSplit(1)), "com") > 0)
    
    If searchSpecificCom Then
        'MsgBox "special closing"
        searchWindowNamePart = strSplit(1) & " - SerialConsole - V1.0 by Ricardo de Roode"
    
        WinWnd = FindWindow(vbNullString, searchWindowNamePart)
    Else
        'MsgBox "global closing"
        searchWindowNamePart = " - SerialConsole - V1.0 by Ricardo de Roode"
        EnumWindows AddressOf EnumWindowsCallBack, 0
        WinWnd = EnumWindowsHWND
    End If
    
    If WinWnd = 0 Then
        'MsgBox "no window"
        Exit Sub
    End If
        
    textboxHandle = FindWindowEx(WinWnd, 0&, "ThunderRT6TextBox", vbNullString)
    
    If textboxHandle = 0 Then
        textboxHandle = FindWindowEx(WinWnd, 0&, "ThunderTextBox", vbNullString)
    End If
    
    If textboxHandle = 0 Then
        'MsgBox "no textbox"
        Exit Sub
    End If
    
    SendMessageByString textboxHandle, WM_SETTEXT, 0&, whatToSend
    
End Sub

Private Function EnumWindowsCallBack(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim sCls As String, lCls As Long
    Dim sCap As String, lCap As Long
    sCls = Space$(255)
    lCls = GetClassName(hWnd, sCls, Len(sCls))
    sCap = Space$(255)
    lCap = GetWindowText(hWnd, sCap, Len(sCap))
    
    If InStr(1, sCap, searchWindowNamePart) > 0 Then
        EnumWindowsHWND = hWnd
    Else
        EnumWindowsCallBack = -1
    End If
End Function

