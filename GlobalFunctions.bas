Attribute VB_Name = "GlobalFunctions"
Option Explicit


Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long

 
Type arduinoLabel
    sKey As String
    lValue As Long
End Type

Sub Main()
    Dim commandProcessed As String
    
    uDontDrawDots = True
    uEnableMouseHooks = True
    
    commandProcessed = Replace(Command, Chr(34), "")
    
    If Len(commandProcessed) > 0 Then
        'MsgBox "'" & commandProcessed & "'"
        SendMessageToInstances commandProcessed
    Else
        Dim t As TypeLibInfo
        Set t = TLI.TypeLibInfoFromFile(App.Path & "/MSCOMM32_ALTERED.OCX")
    
        On Error GoTo NO_DLLS
        
        frmMain.Show
 
        
        'frmTest.Show
    End If
    
    
Exit Sub
NO_DLLS:
    Dim errNum As Long
    errNum = Err.number
    
    If errNum <> 713 And errNum <> 339 Then Exit Sub
    
    If RegSvr32(App.Path & "/MSCOMM32_ALTERED.OCX", False) = True Then
        MsgBox "'MSCOMM32_ALTERED.OCX' was succesfully registered. Please restart the application. If this is not the first time you see this message, please run the application as Administrator.", vbInformation
    Else
        MsgBox "Could not register 'MSCOMM32_ALTERED.OCX'. Please run this program once as Administrator to allow setting up the registry. The program will now close.", vbCritical
    End If
    
    End
End Sub




Public Sub wait(ByVal dblMilliseconds As Double)
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim dblTickCount As Double
    
    dblTickCount = GetTickCount()
    dblStart = GetTickCount()
    dblEnd = GetTickCount + dblMilliseconds
    
    Do
    DoEvents
    dblTickCount = GetTickCount()
    Loop Until dblTickCount > dblEnd Or dblTickCount < dblStart
       
    
End Sub

