Attribute VB_Name = "MultipleInstances"
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const SW_SHOWNORMAL = 1
Private Const WM_CLOSE = &H10
Private Const WM_SETTEXT = &HC


Sub SendMessageToInstances(whatToSend As String)
    Dim WinWnd As Long, Ret As String, RetVal As Long, lpClassName As String
    Dim textboxHandle As Long
    
    Dim strSplit() As String
    strSplit = Split(whatToSend, " ")
    
    If UBound(strSplit) <> 1 Then
        MsgBox "Possible commands:" & vbCrLf & vbCrLf & "CC COM4" & vbCrLf & "OC COM5"
        Exit Sub
    End If
    
    
    
    'Ask for a Window title
    Ret = strSplit(1) & " - SerialConsole - V1.0 by Ricardo de Roode" 'InputBox("Enter the exact window title:" + Chr$(13) + Chr$(10) + "Note: must be an exact match")
    'Search the window
    WinWnd = FindWindow(vbNullString, Ret)
    If WinWnd = 0 Then
        'MsgBox "no window"
        Exit Sub
    End If
    
    textboxHandle = FindWindowEx(WinWnd, 0&, "ThunderRT6TextBox", vbNullString)
    
    If textboxHandle = 0 Then
        'MsgBox "no textbox"
        Exit Sub
    End If
    
    SendMessageByString textboxHandle, WM_SETTEXT, 0&, whatToSend
    
    'Show the window
    'ShowWindow WinWnd, SW_SHOWNORMAL
    'Create a buffer
    'lpClassName = Space(256)
    'retrieve the class name
    'RetVal = GetClassName(WinWnd, lpClassName, 256)
    'Show the classname
    'MsgBox "Classname: " + Left$(lpClassName, RetVal)
    'Post a message to the window to close itself
    'PostMessage WinWnd, WM_CLOSE, 0&, 0&
End Sub
